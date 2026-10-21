# LINE conversation receipt compatibility

Base: production report-api commit 96395f90d19cccfe5a0a55baa9e9953457026501.
The live /health response matched this commit. The live /api/line/config-check confirmed cs-system / public.line_messages; it reports only credential-presence booleans, and no credentials were read or stored.

## Changes

- Persist source.groupId for group chats, roomId for rooms, userId for private chats. Previously userId won even inside a group, splitting one group into several senders. Missing chat IDs remain null; do not guess a common ID.
- Keep the existing user-first *duplicate-hash identity*, timestamp, normalized content and message/event ID inputs for retry compatibility across this release. Historical hashes and message rows are not updated.
- Request PostgREST return=representation and select id,conversation_id. A BEFORE INSERT grouping trigger can suppress a duplicate and return an empty list; count that as duplicate rather than inserted. Invalid response shapes fail rather than claiming success.
- Reject invalid JSON object/event/source/message/text shapes with HTTP 400, after the existing signature validation. Signature, credentials, route authorization and the report API contract are unchanged.
- Text, empty text and media retain their current pending/error statuses and raw content. No cases are created by receipt.

## Shared grouping contract

Both this webhook and the manual importer INSERT the same public.line_messages table. The database trigger owns grouping and aggregates; no second grouping implementation is added to Python.

Companion frontend/SQL draft: repair-form PR #9. Its SQL now accepts unassigned pending **and error** receipts; manual actions can process a pending-review conversation containing mixed pending/error messages while still rejecting linked/archived messages, stale versions and mismatched case relations. This prevents photos from failing receipt after grouping is enabled.

## Validation

- tests/line_ingest_contract.py: 13 standard-library tests of the actual extracted functions. Group/member isolation, rooms/private/missing IDs, unchanged raw text and legacy hashes, redelivery, image status, signature, duplicate precheck/race, failure response and PostgREST headers pass.
- tests/line_http.py: 9 real Flask test-client tests with all external requests mocked. Includes signatures, invalid payloads, group/text/image receipts, duplicate suppression, database failure, and weekly/monthly XLSX export regression.
- Install the existing pinned requirements in an isolated test environment; run both test files with Python. No new production dependencies.
- Companion staging rollback-only SQL tests pass, including mixed text/image grouping and archive, plus the original six grouping scenarios and V3.1 manual case creation. Staging application tables and production data are unchanged.
- Companion synthetic browser harness: 12 checks pass in the existing browser, including raw-text escaping, one conversation per row, history retention and readonly guards. This does not replace authenticated PostgREST/LINE delivery validation.

## Release / rollback

Do not merge directly to production before coordinated review. Deploy this compatible API update first (conversation_id already exists); the new trigger/SQL can then be installed, and the frontend deployed. New conversation materialization starts when the trigger is installed; prior messages are not backfilled.

Verify a newly signed LINE event in a controlled environment, group vs private sender IDs, one retry counted as duplicate, one media event accepted, and no automatic case. Do not send unsigned or fabricated events to production. Exact live LINE channel configuration and signed end-to-end receipt remain unverified.

Rollback API to the base commit if necessary; no database rollback is performed by this repository. Reverting sender identity can split later groups, but does not change already stored raw messages or case relations. Keep the companion SQL capable of pending/error receipts even if the API is reverted.

No customer data, credentials, exports or production mutations are included.
