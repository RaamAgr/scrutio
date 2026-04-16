import { NextRequest, NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import fs from 'fs';
import path from 'path';
import os from 'os';

export const maxDuration = 120;

export async function POST(req: NextRequest) {
  let tmpPath: string | null = null;
  let fileUri: string | null = null;
  let client: GoogleGenAI | null = null;

  try {
    const { fileData, fileName, mimeType, prompt, apiKey, model, temperature, topP, topK, thinkingBudget } = await req.json();

    if (!fileData || !fileName || !prompt || !apiKey) {
      return NextResponse.json({ error: 'Missing required parameters.' }, { status: 400 });
    }

    client = new GoogleGenAI({ apiKey });

    // 1. Decode base64 and write to tmp file
    const buffer = Buffer.from(fileData, 'base64');
    const fileExt = fileName.split('.').pop() || 'mp3';
    tmpPath = path.join(os.tmpdir(), `audio-${Date.now()}-${Math.random().toString(36).slice(2)}.${fileExt}`);
    fs.writeFileSync(tmpPath, buffer);

    console.log(`[evaluate-file] Saved ${fileName} temporarily to ${tmpPath}`);

    // 2. Upload to Gemini Files API
    const resolvedMime = mimeType || `audio/${fileExt === 'wav' ? 'wav' : fileExt === 'ogg' ? 'ogg' : 'mpeg'}`;
    console.log(`[evaluate-file] Uploading ${fileName} to Gemini API...`);
    const uploadResult = await client.files.upload({
      file: tmpPath,
      config: { mimeType: resolvedMime },
    });

    fileUri = uploadResult.name || null;

    // 3. Wait briefly for Gemini to process the file
    await new Promise(resolve => setTimeout(resolve, 3000));

    // 4. Generate content
    console.log(`[evaluate-file] Generating content for ${fileName}...`);
    const genResult = await client.models.generateContent({
      model: model || 'gemini-2.5-flash',
      contents: [
        {
          role: 'user',
          parts: [
            {
              fileData: {
                fileUri: uploadResult.uri,
                mimeType: uploadResult.mimeType,
              },
            },
            { text: prompt },
          ],
        },
      ],
      config: {
        temperature: temperature ? parseFloat(temperature) : 0.7,
        topP: topP ? parseFloat(topP) : 0.95,
        topK: topK ? parseInt(topK) : 64,
        ...(typeof thinkingBudget === 'number' && thinkingBudget >= 0
          ? { thinkingConfig: { thinkingBudget } }
          : {}),
      },
    });

    console.log(`[evaluate-file] Done: ${fileName}`);

    return NextResponse.json({
      success: true,
      result: genResult.text,
      usage: genResult.usageMetadata,
    });
  } catch (error: any) {
    console.error('[evaluate-file] Error:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  } finally {
    if (tmpPath && fs.existsSync(tmpPath)) {
      try { fs.unlinkSync(tmpPath); } catch (e) { /* ignore */ }
    }
    if (client && fileUri) {
      try { await client.files.delete({ name: fileUri }); } catch (e) { /* ignore */ }
    }
  }
}
