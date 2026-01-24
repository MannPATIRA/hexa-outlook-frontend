# Frontend Report: schedule-reply / schedule-replies-batch and supplier_id / supplier_name

**For backend agent: copy-paste the "Summary" and "Details" sections below.**

---

## Summary

- **`/api/demo/schedule-reply`**: The frontend **does** extract `supplier_id` and `supplier_name` from the RFQ response and **does** send them when calling this endpoint, **but only when** it successfully matches the sent email to an RFQ and stores a mapping.
- **`/api/demo/schedule-replies-batch`**: The frontend **never** calls this endpoint. It only uses `schedule-reply` (one request per sent RFQ).
- **Request body keys**: The frontend sends `supplier_id` and `supplier_name` (snake_case) in the JSON body when they are available.

---

## Details

### 1. RFQ response (`POST /api/rfqs/generate`)

The frontend calls `POST /api/rfqs/generate` and stores `response.rfqs` in `AppState.rfqs`. Each RFQ is expected to have:

- `rfq_id`
- `supplier_id`
- `supplier_name`
- `supplier_email`
- `subject`
- (plus other fields)

So the frontend **has** `supplier_id` and `supplier_name` from the RFQ response.

### 2. When are they sent to the backend?

When the user **sends** an RFQ draft:

1. The frontend sends the draft via Graph API, then finds the sent message and gets its `internetMessageId`.
2. It matches that sent email to an RFQ in `AppState.rfqs` by **subject** and **recipient** (`supplier_email`).
3. If a match is found **and** the RFQ has `rfq_id` and `supplier_id`, it stores a mapping:  
   `internetMessageId → { rfq_id, supplier_id, supplier_name, supplier_email }`.
4. It then calls `POST /api/demo/schedule-reply` with:
   - `to_email`, `original_subject`, `original_message_id`, `material`, `reply_type`, `delay_seconds`, `quantity`
   - **and** `supplier_id`, `supplier_name` **when the mapping exists**.

So **when the match succeeds**, the frontend **does** pass `supplier_id` and `supplier_name` to `schedule-reply`.

### 3. When might they be missing?

- **No RFQ match**: If the sent email cannot be matched to an RFQ (e.g. subject/recipient mismatch, or `AppState.rfqs` empty), no mapping is stored. In that case, the frontend does **not** send `supplier_id` or `supplier_name`.
- **RFQ missing `supplier_id`**: The code only stores a mapping when `matchingRfq.rfq_id` and `matchingRfq.supplier_id` exist. If the RFQ from `/api/rfqs/generate` omits `supplier_id`, no mapping is stored and those fields are not sent.
- **Timing / state**: If `AppState.rfqs` is cleared or changed between RFQ generation and send, matching can fail and again no supplier info is sent.

### 4. Exact request body for `POST /api/demo/schedule-reply`

When supplier info **is** included, the frontend sends a JSON body like:

```json
{
  "to_email": "<user email>",
  "original_subject": "<RFQ subject>",
  "original_message_id": "<internetMessageId of sent RFQ>",
  "material": "<material code or 'Unknown Material'>",
  "reply_type": "random",
  "delay_seconds": 5,
  "quantity": 100,
  "supplier_id": "<from RFQ>",
  "supplier_name": "<from RFQ>"
}
```

When **no** mapping exists, `supplier_id` and `supplier_name` are **omitted** (not sent as null).

### 5. `schedule-replies-batch`

The frontend **does not** call `POST /api/demo/schedule-replies-batch`. It only uses `POST /api/demo/schedule-reply`, once per sent RFQ. Any logic or docs referring to `schedule-replies-batch` does not apply to the current frontend.

---

## Backend robustness suggestions

1. **Treat `supplier_id` / `supplier_name` as optional**  
   - Handle missing fields gracefully (e.g. use sender email or “Unknown” when not provided).

2. **`schedule-reply`**  
   - Accept and use `supplier_id` and `supplier_name` when present.  
   - Do not assume they are always present.

3. **`schedule-replies-batch`**  
   - Not used by the frontend. If you support it, treat it as a separate path; the above applies only to `schedule-reply`.

4. **RFQ response**  
   - Ensure `POST /api/rfqs/generate` always returns `supplier_id` (and preferably `supplier_name`) for each RFQ so the frontend can store and forward them.

---

## Code references (frontend)

- RFQ storage: `AppState.rfqs` populated from `ApiClient.generateRFQs()` → `response.rfqs`.
- Mapping storage: `taskpane.js` ~3414–3444 (match by subject + recipient, then `storeRFQMapping`).
- Schedule-reply call: `taskpane.js` ~3474–3516 (`getRFQMapping` → `ApiClient.scheduleAutoReply`).
- API client: `api-client.js` ~217–236 (`scheduleAutoReply` builds payload; adds `supplier_id` / `supplier_name` when provided; `POST /demo/schedule-reply`).
