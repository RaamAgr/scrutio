import { NextRequest, NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';

export const maxDuration = 60; // Vercel higher timeout if deployed

export async function POST(req: NextRequest) {
  try {
    const { feed, prompt, apiKey, model, temperature } = await req.json();

    if (!feed || !prompt || !apiKey) {
      return NextResponse.json({ error: 'Missing required parameters: feed, prompt, or apiKey.' }, { status: 400 });
    }

    const client = new GoogleGenAI({ apiKey });

    console.log(`Starting Level 2 Meta-Analysis using model ${model || 'gemini-2.5-flash'}...`);

    const fullPrompt = `You are a high-level data synthesis and meta-analysis engine. 
Below is a massive transcript of results gathered from a first-pass evaluation on several calls.

User's analytical instructions:
---
${prompt}
---

Here is the raw data feed:
=== START OF FEED ===
${feed}
=== END OF FEED ===

Please execute the analytical instructions precisely.`;

    const genResult = await client.models.generateContent({
      model: model || 'gemini-2.5-flash',
      contents: fullPrompt,
      config: {
        temperature: temperature ? parseFloat(temperature) : 0.4,
      }
    });

    console.log('Meta-Evaluation complete.');

    return NextResponse.json({
      success: true,
      result: genResult.text,
      usage: genResult.usageMetadata
    });

  } catch (error: any) {
    console.error('Error during meta-evaluation:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
