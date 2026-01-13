# Backend Question Parsing - Implementation Prompt

## Context

The Outlook Procurement Add-in frontend has been updated to parse and display clarification questions as separate, individual questions instead of a single text block. The frontend currently handles this parsing client-side, but we want to confirm if backend changes would be beneficial.

## Current Frontend Implementation

The frontend now:
- Parses questions from email body text using pattern matching
- Extracts numbered sections, bullet points, and question formats
- Groups questions by category
- Displays each question as a separate, numbered item

## Questions for Backend Team

### 1. Does the backend already extract questions during email classification?

**Current API:** `POST /api/emails/classify`

**Question:** When classifying a `clarification_request`, does the backend already extract individual questions from the email body? If so, are they returned in the classification response?

**If Yes:** Please provide the response structure so we can update the frontend to use backend-parsed questions instead of client-side parsing.

**If No:** Would it be beneficial to add question extraction to the classification endpoint?

### 2. Should we add a dedicated question extraction endpoint?

**Proposed Endpoint:** `POST /api/emails/extract-questions`

**Request:**
```json
{
  "email_body": "Full email body text or HTML",
  "email_id": "optional - for context"
}
```

**Response:**
```json
{
  "questions": [
    {
      "category": "Tolerance Requirements",
      "question": "What are the critical dimension tolerances?",
      "section_number": "1",
      "confidence": 0.95
    },
    {
      "category": "Tolerance Requirements",
      "question": "Are there any specific GD&T callouts we should be aware of?",
      "section_number": "1",
      "confidence": 0.92
    }
  ],
  "parsing_method": "nlp" | "pattern_matching" | "hybrid"
}
```

**Benefits:**
- More accurate parsing using NLP/AI models
- Consistent extraction across all clients
- Better handling of edge cases
- Ability to link questions to RFQ requirements

### 3. Should question extraction be part of the `/api/emails/process` endpoint?

**Current API:** `POST /api/emails/process`

**Question:** When processing a clarification email, should the backend return structured question data along with the `suggested_response` and `requires_engineering` fields?

**Proposed Response Enhancement:**
```json
{
  "clarification_id": "...",
  "sub_classification": "procurement" | "engineering",
  "requires_engineering": false,
  "suggested_response": "...",
  "questions": [
    {
      "category": "...",
      "question": "...",
      "section_number": "..."
    }
  ]
}
```

### 4. Integration with AI/LLM Models

**Question:** If the backend uses AI/LLM models for email classification, could these models also extract individual questions more accurately than pattern matching?

**Considerations:**
- LLM models can understand context better
- Can handle various email formats and languages
- Can identify question intent even without explicit question marks
- More robust than regex/pattern matching

### 5. Question Linking to RFQ Requirements

**Question:** Should extracted questions be linked to specific RFQ line items or requirements?

**Use Case:** If a supplier asks "What are the critical dimension tolerances?", the backend could:
- Identify which RFQ requirement this relates to
- Provide context-aware suggested responses
- Track which requirements need clarification

## Recommended Approach

### Option 1: Enhance Existing Endpoints (Recommended)
- Add question extraction to `/api/emails/classify` or `/api/emails/process`
- Return structured questions in the response
- Frontend will use backend questions if available, fallback to client-side parsing

### Option 2: New Dedicated Endpoint
- Create `/api/emails/extract-questions` endpoint
- Frontend calls this after classification
- Allows for independent question extraction service

### Option 3: No Backend Changes (Current State)
- Frontend continues to parse questions client-side
- Backend focuses on classification and response generation
- Simpler architecture, but less accurate parsing

## Implementation Notes

1. **Backward Compatibility:** If backend provides questions, frontend will use them. If not, frontend falls back to client-side parsing.

2. **Question Format:** Questions should include:
   - `category` (optional): Group/category name
   - `question`: The actual question text
   - `section_number` (optional): Original section number from email
   - `confidence` (optional): Parsing confidence score

3. **Error Handling:** If question extraction fails on backend, frontend should gracefully fall back to client-side parsing.

## Next Steps

Please confirm:
1. ✅ Does backend already extract questions?
2. ✅ Should we add question extraction to existing endpoints?
3. ✅ Should we create a new dedicated endpoint?
4. ✅ What's the recommended approach?

Once confirmed, we can:
- Update frontend to consume backend-provided questions
- Remove or keep client-side parsing as fallback
- Update API documentation

## Contact

For questions or clarifications about the frontend implementation, please refer to `QUESTION_PARSING_IMPLEMENTATION.md`.
