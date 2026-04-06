import { NextRequest, NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import fs from 'fs';
import path from 'path';
import os from 'os';

export const maxDuration = 60; // Vercel higher timeout if deployed

export async function POST(req: NextRequest) {
  let tmpPath: string | null = null;
  let fileUri: string | null = null;
  let client: GoogleGenAI | null = null;

  try {
    const { url, prompt, apiKey, model, temperature, topP, topK, thinkingBudget } = await req.json();

    if (!url || !prompt || !apiKey) {
      return NextResponse.json({ error: 'Missing required parameters: url, prompt, or apiKey.' }, { status: 400 });
    }

    client = new GoogleGenAI({ apiKey });

    // 1. Download the audio file
    console.log(`Downloading audio from ${url}`);
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Failed to download audio. Status: ${response.status} ${response.statusText}`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    // 2. Save it to a temporary file to upload via SDK
    // @google/genai currently prefers paths for file uploads, or it handles it better.
    const fileExt = url.split('.').pop()?.split('?')[0] || 'mp3';
    tmpPath = path.join(os.tmpdir(), `audio-${Date.now()}.${fileExt}`);
    fs.writeFileSync(tmpPath, buffer);

    console.log(`Audio saved temporarily to ${tmpPath}`);

    // 3. Upload to Gemini
    console.log('Uploading file to Gemini API...');
    const uploadResult = await client.files.upload({
      file: tmpPath,
      config: {
        mimeType: response.headers.get('content-type') || `audio/${fileExt === 'wav' ? 'wav' : 'mpeg'}`,
      }
    });
    
    fileUri = uploadResult.name || null; // The unique identifier returned

    console.log('Waiting slightly for file processing...');
    // Sometimes the file needs a moment to be PROCESSED by Gemini
    await new Promise(resolve => setTimeout(resolve, 3000));

    // 4. Generate Content
    console.log(`Generating content using model ${model || 'gemini-2.5-flash'}...`);
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
              }
            },
            { text: prompt }
          ]
        }
      ],
      config: {
        temperature: temperature ? parseFloat(temperature) : 0.7,
        topP: topP ? parseFloat(topP) : 0.95,
        topK: topK ? parseInt(topK) : 64,
        ...(typeof thinkingBudget === 'number' && thinkingBudget >= 0
          ? { thinkingConfig: { thinkingBudget } }
          : {}),
      }
    });

    console.log('Evaluation complete.');

    return NextResponse.json({
      success: true,
      result: genResult.text,
      usage: genResult.usageMetadata
    });

  } catch (error: any) {
    console.error('Error during evaluation:', error);
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  } finally {
    // 5. Cleanup
    if (tmpPath && fs.existsSync(tmpPath)) {
      try {
        fs.unlinkSync(tmpPath);
      } catch (e) {
        console.error('Failed to clear tmp file:', e);
      }
    }
    if (client && fileUri) {
      try {
        await client.files.delete({ name: fileUri });
      } catch (e) {
        console.error('Failed to remote file from Gemini:', e);
      }
    }
  }
}
