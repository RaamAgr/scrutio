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

    const isVertex = apiKey.startsWith('AQ.');
    client = new GoogleGenAI(isVertex ? { apiKey, vertexai: true } : { apiKey });

    // 1. Download the audio file
    console.log(`Downloading audio from ${url}`);
    const response = await fetch(url);
    if (!response.ok) {
      throw new Error(`Failed to download audio. Status: ${response.status} ${response.statusText}`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const fileExt = url.split('.').pop()?.split('?')[0] || 'mp3';
    const mimeType = response.headers.get('content-type') || `audio/${fileExt === 'wav' ? 'wav' : 'mpeg'}`;

    let genResult;

    if (isVertex) {
      console.log('Vertex AI detected (starts with AQ.). Using inlineData (base64) instead of Files API.');
      genResult = await client.models.generateContent({
        model: model || 'gemini-3-flash-preview',
        contents: [
          {
            role: 'user',
            parts: [
              {
                inlineData: {
                  data: buffer.toString('base64'),
                  mimeType: mimeType,
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
    } else {
      // 2. Save it to a temporary file to upload via SDK for AI Studio
      tmpPath = path.join(os.tmpdir(), `audio-${Date.now()}.${fileExt}`);
      fs.writeFileSync(tmpPath, buffer);

      console.log(`Audio saved temporarily to ${tmpPath}`);

      // 3. Upload to Gemini
      console.log('Uploading file to Gemini API...');
      const uploadResult = await client.files.upload({
        file: tmpPath,
        config: { mimeType }
      });
      
      fileUri = uploadResult.name || null; // The unique identifier returned

      console.log('Waiting slightly for file processing...');
      // Sometimes the file needs a moment to be PROCESSED by Gemini
      await new Promise(resolve => setTimeout(resolve, 3000));

      // 4. Generate Content
      console.log(`Generating content using model ${model || 'gemini-2.5-flash'}...`);
      genResult = await client.models.generateContent({
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
    }

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
