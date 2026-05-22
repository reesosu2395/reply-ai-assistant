export default {
  async fetch(request, env) {
    // CORS preflight
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders() });
    }

    if (request.method !== "POST") {
      return json({ error: "Method not allowed" }, 405);
    }

    const url = new URL(request.url);
    if (url.pathname !== "/get-replies") {
      return json({ error: "Not found" }, 404);
    }

    let emailBody;
    try {
      const body = await request.json();
      emailBody = body.emailBody?.trim();
      if (!emailBody) throw new Error("Empty body");
    } catch {
      return json({ error: "Invalid request body" }, 400);
    }

    const prompt = buildPrompt(emailBody);

    try {
      const aiResponse = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": env.ANTHROPIC_API_KEY,       // Set in Cloudflare dashboard
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1024,
          system: `You are an email assistant for a coatings company sales/support team.
                   Generate exactly 4 distinct reply options. Return ONLY valid JSON in
                   this format: {"options": ["reply1", "reply2", "reply3", "reply4"]}
                   No markdown, no extra text.`,
          messages: [{ role: "user", content: prompt }],
        }),
      });

      const result = await aiResponse.json();
      const text = result.content?.[0]?.text ?? "{}";
      const parsed = JSON.parse(text);

      if (!Array.isArray(parsed.options) || parsed.options.length !== 4) {
        throw new Error("Unexpected AI response shape");
      }

      return json(parsed);
    } catch (err) {
      console.error("AI call failed:", err);
      return json({ error: "AI service error", detail: err.message }, 502);
    }
  },
};

// --- Helpers ---

function buildPrompt(emailBody) {
  return `
Email received:
"""
${emailBody}
"""

Generate 4 distinct professional reply options for a Gemini Coatings team member.
Each reply should differ in tone or approach:
1. Formal & detailed
2. Friendly & concise
3. Action-oriented (next steps focused)
4. Empathetic / relationship-building

Return ONLY the JSON object with the "options" array.
  `.trim();
}

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",        // Restrict to your add-in domain in prod
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

function json(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json", ...corsHeaders() },
  });
}