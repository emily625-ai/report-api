"""Real Flask test-client tests with all external requests mocked."""
import base64, hashlib, hmac, importlib.util, json, os
from pathlib import Path
from unittest import TestCase, main
from unittest.mock import Mock, patch

os.environ.update(LINE_MESSAGES_SCHEMA='public', LINE_CHANNEL_SECRET='synthetic-secret')
spec=importlib.util.spec_from_file_location('line_test_app',Path(__file__).resolve().parents[1]/'app.py')
module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)

class LineHTTP(TestCase):
 def setUp(self):
  self.client=module.app.test_client()
  self.outbound=patch.object(module.requests,'request',side_effect=AssertionError('Unexpected external request'));self.outbound.start();self.addCleanup(self.outbound.stop)
  self.profiles=patch.object(module,'line_api_get',return_value=None);self.profiles.start();self.addCleanup(self.profiles.stop)
  self.rest=patch.object(module,'supabase_request',side_effect=self.database);self.mock_rest=self.rest.start();self.addCleanup(self.rest.stop)
  self.inserts=[]
 def database(self,method,table,**kwargs):
  if method=='GET':return []
  if table=='activity_log':return None
  self.inserts.append((table,kwargs.get('payload')))
  return [{'id':'synthetic-id','conversation_id':'synthetic-conversation'}]
 def send(self,payload,signature=None):
  body=json.dumps(payload).encode()
  signature=signature or base64.b64encode(hmac.new(b'synthetic-secret',body,hashlib.sha256).digest()).decode()
  return self.client.post('/api/line/webhook',data=body,headers={'x-line-signature':signature,'Content-Type':'application/json'})
 def event(self,typ='text'):
  return {'type':'message','timestamp':1792540800000,'webhookEventId':'synthetic-event','source':{'type':'group','groupId':'G','userId':'U'},'message':{'type':typ,'id':'synthetic-message','text':'synthetic original'}}
 def test_invalid_signature_denied_before_database(self):
  self.assertEqual(self.send({'events':[self.event()]},'wrong').status_code,403);self.mock_rest.assert_not_called()
 def test_empty_verification_request(self):
  r=self.send({'events':[]});self.assertEqual(r.status_code,200);self.assertEqual(r.json['received'],0);self.mock_rest.assert_not_called()
 def test_malformed_payloads(self):
  for value in [[],{'events':'bad'},{'events':[None]},{'events':[{'type':'message','source':['bad']}]},{'events':[{'type':'message','message':{'type':'text','text':42}}]}]:
   with self.subTest(payload=value):self.assertEqual(self.send(value).status_code,400)
  self.mock_rest.assert_not_called()
 def test_invalid_json(self):
  body=b'not json';signature=base64.b64encode(hmac.new(b'synthetic-secret',body,hashlib.sha256).digest()).decode()
  self.assertEqual(self.client.post('/api/line/webhook',data=body,headers={'x-line-signature':signature}).status_code,400)
 def test_text_group_identity_and_no_case_write(self):
  r=self.send({'events':[self.event()]});self.assertEqual(r.status_code,200);self.assertEqual(r.json['inserted'],1)
  self.assertEqual(self.inserts[0][0],'line_messages');self.assertEqual(self.inserts[0][1]['sender_id'],'G')
  self.assertEqual(self.inserts[0][1]['raw_message'],'synthetic original');self.assertNotIn('case_id',self.inserts[0][1]);self.assertEqual(len(self.inserts),1)
 def test_image_receipt(self):
  r=self.send({'events':[self.event('image')]});self.assertEqual(r.status_code,200);self.assertEqual(self.inserts[0][1]['status'],'error')
 def test_atomic_duplicate_suppression(self):
  self.mock_rest.side_effect=lambda method,table,**kwargs: [] if table=='line_messages' else None
  r=self.send({'events':[self.event()]});self.assertEqual(r.status_code,200);self.assertEqual(r.json['duplicates'],1);self.assertEqual(r.json['inserted'],0)
 def test_database_unavailable_returns_failure(self):
  self.mock_rest.side_effect=RuntimeError('synthetic database unavailable')
  with patch.object(module.app.logger,'exception'),patch.object(module.app.logger,'warning'):
   self.assertEqual(self.send({'events':[self.event()]}).status_code,500)
 def test_weekly_and_monthly_export_regression(self):
  rows=[{'id':'synthetic-case','date':'2026-10-21 08:00','company':'Synthetic','product':'FMS-GPS(系統平台)','category':'平台系統','subcategory':'定位異常','status':'客服處理中','new_product':'GPS','case_nature':'故障/異常'}]
  for endpoint in ['/weekly-report','/monthly-report']:
   with self.subTest(endpoint=endpoint):
    r=self.client.post(endpoint,json={'records':rows,'all_records':rows,'from':'2026-10-01','to':'2026-10-31'})
    self.assertEqual(r.status_code,200,r.get_data(as_text=False)[:200]);self.assertTrue(r.data.startswith(b'PK'))

if __name__=='__main__':main()
