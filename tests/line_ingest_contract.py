"""Standard-library contract tests on the real webhook functions extracted by AST.
No production credentials, LINE calls, database calls or Flask installation required.
"""
import ast, base64, copy, hashlib, hmac, json, logging, os
from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from unittest import TestCase, main
from unittest.mock import Mock, patch

root = Path(__file__).resolve().parents[1]
source = (root / 'app.py').read_text()
functions = {'get_required_env','normalize_line_message','verify_line_signature','line_timestamp_to_iso',
 'resolve_line_sender_name','build_line_duplicate_hash','get_supabase_headers','supabase_rest_url',
 'supabase_request','line_message_exists','insert_line_message','log_line_activity','build_line_message_row',
 'summarize_line_event_for_log','process_line_webhook_events','line_webhook'}
nodes = []
for node in ast.parse(source).body:
 if isinstance(node, ast.FunctionDef) and node.name in functions:
  node.decorator_list = []
  nodes.append(node)
code = compile(ast.Module(body=nodes,type_ignores=[]), 'app.py:webhook', 'exec')

class Ingest(TestCase):
 def setUp(self):
  self.env = patch.dict(os.environ, {'LINE_CHANNEL_SECRET':'synthetic-test-secret','SUPABASE_SERVICE_ROLE_KEY':'synthetic-key','SUPABASE_URL':'https://synthetic.invalid'}, clear=True)
  self.env.start();self.addCleanup(self.env.stop)
  self.ns = dict(os=os,json=json,hashlib=hashlib,hmac=hmac,base64=base64,datetime=datetime,timezone=timezone,
   requests=SimpleNamespace(request=Mock()),app=SimpleNamespace(logger=Mock()),
   LINE_WEBHOOK_SOURCE='line_webhook',LINE_MESSAGES_TABLE='line_messages',LINE_MESSAGES_SCHEMA='public')
  exec(code,self.ns)
  self.ns['line_api_get']=lambda path: None
  self.ns['log_line_activity']=Mock()
 def event(self, **kwargs):
  event={'type':'message','webhookEventId':'synthetic-event-1','timestamp':1792540800000,
   'source':{'type':'group','groupId':'synthetic-G','userId':'synthetic-U'},
   'message':{'type':'text','id':'synthetic-message-1','text':'  Synthetic original\nmessage  '}}
  event.update(kwargs);return event
 def test_same_group_different_members(self):
  a=self.ns['build_line_message_row'](self.event());e=self.event();e['source']['userId']='synthetic-U2'
  b=self.ns['build_line_message_row'](e)
  self.assertEqual(a['sender_id'],b['sender_id']);self.assertEqual(a['sender_id'],'synthetic-G')
 def test_different_groups_same_member(self):
  a=self.ns['build_line_message_row'](self.event());e=self.event();e['source']['groupId']='synthetic-G2'
  self.assertNotEqual(a['sender_id'],self.ns['build_line_message_row'](e)['sender_id'])
 def test_room_private_and_missing(self):
  for src,expected in [({'type':'room','roomId':'R','userId':'U'},'R'),({'type':'user','userId':'U'},'U'),({'type':'group','userId':'U'},None),({},None)]:
   self.assertEqual(self.ns['build_line_message_row'](self.event(source=src))['sender_id'],expected)
 def test_raw_message_unchanged_no_case(self):
  e=self.event();before=copy.deepcopy(e);row=self.ns['build_line_message_row'](e)
  self.assertEqual(row['raw_message'],e['message']['text']);self.assertEqual(e,before)
  self.assertNotIn('case_id',row);self.assertNotIn('conversation_id',row)
 def test_existing_hash_stable_across_rollout(self):
  e=self.event();row=self.ns['build_line_message_row'](e)
  old=self.ns['build_line_duplicate_hash']('synthetic-U','LINE 群組',row['received_at'],row['normalized_message'],e['message']['id'])
  self.assertEqual(row['duplicate_hash'],old)
  e['deliveryContext']={'isRedelivery':True}
  self.assertEqual(self.ns['build_line_message_row'](e)['duplicate_hash'],old)
 def test_media_preserves_error_status(self):
  row=self.ns['build_line_message_row'](self.event(message={'id':'media-1','type':'image'}))
  self.assertEqual(row['status'],'error');self.assertEqual(json.loads(row['raw_message'])['message_type'],'image')
 def test_signature_and_invalid_payload(self):
  body=json.dumps({'events':[]}).encode();signature=base64.b64encode(hmac.new(b'synthetic-test-secret',body,hashlib.sha256).digest()).decode()
  self.assertTrue(self.ns['verify_line_signature'](body,signature));self.assertFalse(self.ns['verify_line_signature'](body,'bad'))
  self.ns['request']=SimpleNamespace(get_data=lambda:body,headers={'x-line-signature':'bad'})
  self.ns['jsonify']=lambda value:value
  self.assertEqual(self.ns['line_webhook']()[1],403)
  self.ns['request'].headers={'x-line-signature':signature}
  self.assertEqual(self.ns['line_webhook']()['inserted'],0)
 def run_events(self,exists,inserted):
  self.ns['line_message_exists']=Mock(return_value=exists)
  self.ns['insert_line_message']=Mock(return_value=inserted)
  return self.ns['process_line_webhook_events']([self.event()])
 def test_success_and_no_auto_case(self):
  result=self.run_events(False,[{'id':'synthetic','conversation_id':'synthetic-conversation'}])
  self.assertEqual(result['inserted'],1);self.assertEqual(result['duplicates'],0)
 def test_precheck_duplicate(self):
  result=self.run_events(True,None);self.assertEqual(result['duplicates'],1);self.ns['insert_line_message'].assert_not_called()
 def test_race_duplicate_returning_empty(self):
  result=self.run_events(False,[]);self.assertEqual(result['duplicates'],1);self.assertEqual(result['inserted'],0)
 def test_invalid_insert_response_fails(self):
  with self.assertRaises(RuntimeError):self.run_events(False,None)
 def test_database_error_not_reported_as_success(self):
  self.ns['line_message_exists']=lambda _:False
  self.ns['insert_line_message']=Mock(side_effect=RuntimeError('synthetic outage'))
  with self.assertRaises(RuntimeError):self.ns['process_line_webhook_events']([self.event()])
 def test_postgrest_contract(self):
  response=SimpleNamespace(status_code=201,text='[{"id":"synthetic","conversation_id":"conversation"}]',json=lambda:[{'id':'synthetic','conversation_id':'conversation'}])
  self.ns['requests'].request.return_value=response
  rows=self.ns['insert_line_message']({'raw_message':'synthetic'})
  args,kw=self.ns['requests'].request.call_args
  self.assertEqual(args[0],'POST');self.assertIn('select=id,conversation_id',args[1])
  self.assertEqual(kw['headers']['Prefer'],'return=representation');self.assertEqual(kw['headers']['Content-Profile'],'public')
  self.assertEqual(rows[0]['conversation_id'],'conversation')

if __name__=='__main__':main()
