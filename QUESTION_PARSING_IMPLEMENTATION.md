# Question Parsing Implementation - Frontend Changes

## Overview
This document explains how clarification questions are now parsed and displayed as separate questions in the Outlook add-in, instead of showing them as a single text block.

## Frontend Implementation

### How It Works

1. **Question Parsing Function** (`src/utils/helpers.js`)
   - Added `parseClarificationQuestions(emailBody)` function
   - Extracts individual questions from email body text
   - Supports multiple formats:
     - Numbered sections with bullet points (e.g., "1. Tolerance Requirements\n- Question 1\n- Question 2")
     - Simple bullet points without numbered sections
     - Questions ending with "?" or starting with question words
     - Fallback: splitting by common separators

2. **HTML Structure** (`src/taskpane/taskpane.html`)
   - Updated clarification question section to include:
     - `clarification-questions-list` - Container for individual question items
     - `clarification-question-text` - Fallback single text box (hidden by default)

3. **Display Logic** (`src/taskpane/taskpane.js`)
   - Modified `showClarificationMode()` function
   - Parses email body using `Helpers.parseClarificationQuestions()`
   - Groups questions by category (if available)
   - Renders each question as a separate item with:
     - Question number
     - Question text
     - Category header (if multiple categories exist)

4. **Styling** (`src/taskpane/taskpane.css`)
   - Added styles for:
     - `.questions-list` - Container for question items
     - `.question-item` - Individual question card with hover effect
     - `.question-number` - Numbered indicator
     - `.question-text` - Question content
     - `.question-category-header` - Category grouping header

### Example Output

**Before:**
```
Supplier's Question
[Single large text block with all questions]
```

**After:**
```
Supplier's Questions

Tolerance Requirements
1. What are the critical dimension tolerances?
2. Are there any specific GD&T callouts we should be aware of?

Material Specifications
3. Please confirm the exact material grade (e.g., ASTM standard)
4. Are there specific hardness requirements?
```

## Backend Considerations

### Current Frontend Behavior
- The frontend parses questions **client-side** from the email body text
- No backend API changes are required for basic functionality
- Questions are extracted from `email.body.content` that comes from Microsoft Graph API

### Potential Backend Enhancements

The backend could optionally provide structured question data to improve accuracy and consistency:

1. **Enhanced Classification Response**
   - If the backend already extracts questions during classification, it could return them in a structured format:
   ```json
   {
     "classification": "clarification_request",
     "questions": [
       {
         "category": "Tolerance Requirements",
         "question": "What are the critical dimension tolerances?",
         "section_number": "1"
       },
       {
         "category": "Tolerance Requirements",
         "question": "Are there any specific GD&T callouts we should be aware of?",
         "section_number": "1"
       }
     ]
   }
   ```

2. **Question Extraction Endpoint**
   - A dedicated endpoint to extract questions from email body:
   ```
   POST /api/emails/extract-questions
   {
     "email_body": "...",
     "email_id": "..."
   }
   ```

3. **Benefits of Backend Parsing**
   - More accurate parsing using NLP/AI models
   - Consistent question extraction across all clients
   - Better handling of edge cases and various email formats
   - Ability to link questions to specific RFQ line items or requirements

### Recommendation

**Current State:** Frontend parsing works independently and doesn't require backend changes.

**Future Enhancement:** If the backend team wants to provide structured question data, the frontend can be updated to:
1. Prefer backend-provided questions if available
2. Fall back to frontend parsing if backend doesn't provide questions
3. Merge/validate frontend-parsed questions with backend-provided questions

## Testing

To test the question parsing:

1. Open a clarification email in Outlook
2. The add-in should automatically parse and display questions separately
3. Questions should be grouped by category (if applicable)
4. Each question should appear as a numbered item

## Files Modified

- `src/utils/helpers.js` - Added `parseClarificationQuestions()` function
- `src/taskpane/taskpane.html` - Updated HTML structure
- `src/taskpane/taskpane.js` - Updated display logic
- `src/taskpane/taskpane.css` - Added question list styles
