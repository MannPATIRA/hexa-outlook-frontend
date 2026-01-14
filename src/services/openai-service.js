/**
 * OpenAI Service
 * Handles question parsing and response generation using OpenAI API
 * 
 * SECURITY WARNING: API key is stored in frontend code
 * In production, API calls should be proxied through backend
 */
const OpenAIService = {
    /**
     * Make a request to OpenAI API
     */
    async request(endpoint, options = {}) {
        // Check if API key is configured
        if (!Config.OPENAI_API_KEY) {
            throw new Error('OpenAI API key is not configured. Please set OPENAI_API_KEY in Vercel environment variables or localStorage.');
        }
        
        const url = `${Config.OPENAI_API_BASE_URL}${endpoint}`;
        
        const defaultOptions = {
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${Config.OPENAI_API_KEY}`
            },
        };

        const fetchOptions = { ...defaultOptions, ...options };
        
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), Config.REQUEST_TIMEOUT);
            
            const response = await fetch(url, {
                ...fetchOptions,
                signal: controller.signal
            });
            
            clearTimeout(timeoutId);

            if (!response.ok) {
                const errorData = await response.json().catch(() => ({}));
                // Provide helpful error messages for common issues
                if (response.status === 401) {
                    throw new Error('OpenAI API key is invalid or expired. Please check your OPENAI_API_KEY configuration.');
                } else if (response.status === 429) {
                    throw new Error('OpenAI API rate limit exceeded. Please try again later.');
                } else if (response.status === 500 || response.status === 502 || response.status === 503) {
                    throw new Error('OpenAI API service is temporarily unavailable. Please try again later.');
                }
                throw new Error(
                    errorData.error?.message || `HTTP ${response.status}: ${response.statusText}`
                );
            }

            return await response.json();
        } catch (error) {
            if (error.name === 'AbortError') {
                throw new Error('Request timed out');
            }
            throw error;
        }
    },

    /**
     * Parse questions from email body using OpenAI
     * @param {string} emailBody - The email body text (HTML or plain text)
     * @param {string} emailSubject - Optional email subject for context
     * @returns {Promise<Array>} Array of question objects with {category, question, section_number}
     */
    async parseQuestions(emailBody, emailSubject = '') {
        try {
            // Convert HTML to plain text if needed
            const text = Helpers.stripHtml(emailBody);
            
            const prompt = `You are analyzing an email from a supplier asking clarification questions about a procurement request.

Email Subject: ${emailSubject || 'Not provided'}

Email Body:
${text}

Please extract all individual questions from this email and return them as a JSON array. Each question should have:
- "category": A brief category name (e.g., "Tolerance Requirements", "Material Specifications")
- "question": The exact question text
- "section_number": The section number if the question is part of a numbered list (null if not)

Return ONLY a valid JSON array, no other text. Example format:
[
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
]`;

            const response = await this.request('/chat/completions', {
                method: 'POST',
                body: JSON.stringify({
                    model: 'gpt-4o-mini',
                    messages: [
                        {
                            role: 'system',
                            content: 'You are a helpful assistant that extracts questions from emails. Always return valid JSON arrays only.'
                        },
                        {
                            role: 'user',
                            content: prompt
                        }
                    ],
                    temperature: 0.3,
                    max_tokens: 2000
                })
            });

            // Extract JSON from response
            const content = response.choices[0]?.message?.content || '[]';
            
            // Try to parse JSON (handle cases where response might have markdown code blocks)
            let questions = [];
            try {
                // Remove markdown code blocks if present
                const jsonMatch = content.match(/\[[\s\S]*\]/);
                const jsonText = jsonMatch ? jsonMatch[0] : content;
                questions = JSON.parse(jsonText);
            } catch (parseError) {
                console.error('Failed to parse OpenAI response as JSON:', parseError);
                throw new Error('Failed to parse questions from AI response');
            }

            // Validate and clean questions
            if (!Array.isArray(questions)) {
                throw new Error('Invalid response format: expected array');
            }

            return questions
                .filter(q => q.question && q.question.trim().length > 0)
                .map(q => ({
                    category: q.category || 'General Questions',
                    question: q.question.trim(),
                    section_number: q.section_number || null
                }));
        } catch (error) {
            console.error('OpenAI parseQuestions error:', error);
            throw error;
        }
    },

    /**
     * Generate AI response for a single question
     * @param {string} question - The question to answer
     * @param {Object} emailContext - Context about the email (subject, body, etc.)
     * @returns {Promise<string>} AI-generated response
     */
    async generateResponse(question, emailContext = {}) {
        try {
            const { subject = '', body = '', rfqContext = '' } = emailContext;
            const emailText = Helpers.stripHtml(body).substring(0, 2000); // Limit context length

            const prompt = `You are a procurement specialist responding to a supplier's clarification question.

Email Subject: ${subject || 'Not provided'}

Email Context (relevant parts):
${emailText}

RFQ Context:
${rfqContext || 'Not provided'}

Supplier's Question:
${question}

Please provide a direct, concise answer to this question. The response should:
1. Answer the question directly - no greetings, no closings, no email formatting
2. Be clear and professional
3. Be just the factual answer to the question
4. If you don't have enough information, state what information is needed

IMPORTANT: Provide ONLY the answer to the question. Do NOT include:
- Greetings like "Dear" or "Thank you"
- Closings like "Best regards" or "Sincerely"
- Email formatting
- References to "we" or "our team" unless necessary for the answer
- Just provide the direct answer to the question.`;

            const response = await this.request('/chat/completions', {
                method: 'POST',
                body: JSON.stringify({
                    model: 'gpt-4o-mini',
                    messages: [
                        {
                            role: 'system',
                            content: 'You are a professional procurement specialist. Provide clear, helpful responses to supplier questions.'
                        },
                        {
                            role: 'user',
                            content: prompt
                        }
                    ],
                    temperature: 0.7,
                    max_tokens: 500
                })
            });

            const aiResponse = response.choices[0]?.message?.content || '';
            return aiResponse.trim();
        } catch (error) {
            console.error('OpenAI generateResponse error:', error);
            throw error;
        }
    }
};
