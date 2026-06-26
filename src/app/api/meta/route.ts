import { NextRequest, NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import { GoogleAuth } from 'google-auth-library';

export const maxDuration = 60; // Vercel higher timeout if deployed

export async function POST(req: NextRequest) {
  try {
    const { feed, prompt, apiKey, anthropicApiKey, vertexServiceAccount, model, temperature } = await req.json();

    if (!feed || !prompt) {
      return NextResponse.json({ error: 'Missing required parameters: feed or prompt.' }, { status: 400 });
    }

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

    if (model === 'claude-sonnet-4-6') {
      const activeKey = anthropicApiKey || apiKey;

      // Path A: Direct Anthropic API Key
      if (activeKey && activeKey.startsWith('sk-')) {
        console.log('Active key is an Anthropic Key. Routing Level 2 Meta-Analysis directly to Anthropic API...');
        
        const response = await fetch('https://api.anthropic.com/v1/messages', {
          method: 'POST',
          headers: {
            'x-api-key': activeKey,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json'
          },
          body: JSON.stringify({
            model: 'claude-3-5-sonnet-20241022',
            max_tokens: 4096,
            temperature: temperature ? parseFloat(temperature) : 0.4,
            messages: [
              {
                role: 'user',
                content: fullPrompt
              }
            ]
          })
        });

        if (!response.ok) {
          const errText = await response.text();
          throw new Error(`Anthropic API error: ${response.status} ${response.statusText} - ${errText}`);
        }

        const data = await response.json();
        console.log('Anthropic Meta-Evaluation complete.');

        return NextResponse.json({
          success: true,
          result: data.content[0].text,
          usage: {
            promptTokenCount: data.usage.input_tokens,
            candidatesTokenCount: data.usage.output_tokens
          }
        });
      }

      // Path B: Google Cloud Vertex AI via Application Default Credentials (ADC) or Service Account
      console.log('Routing Level 2 Meta-Analysis to Google Vertex AI using local credentials...');
      try {
        const authOptions: any = {
          scopes: 'https://www.googleapis.com/auth/cloud-platform'
        };

        if (vertexServiceAccount) {
          console.log('Using user-provided Vertex Service Account JSON key...');
          authOptions.credentials = JSON.parse(vertexServiceAccount);
        }

        const auth = new GoogleAuth(authOptions);
        
        const projectId = await auth.getProjectId();
        const client = await auth.getClient();
        const tokenResponse = await client.getAccessToken();
        const accessToken = tokenResponse.token;

        if (!projectId || !accessToken) {
          throw new Error('Unable to resolve GCP Project ID or Access Token from the environment.');
        }

        const candidateModels = [
          'claude-3-5-sonnet-v2@20241022',
          'claude-3-5-sonnet@20240620',
          'claude-3-5-sonnet',
          'claude-3-5-sonnet-v2'
        ];
        
        let lastError = null;
        let data = null;
        
        const regions = ['us-east5', 'us-central1', 'europe-west3'];
        
        for (const region of regions) {
          for (const modelId of candidateModels) {
            const url = `https://${region}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${region}/publishers/anthropic/models/${modelId}:predict`;
            console.log(`Attempting Vertex AI prediction on region ${region} with model: ${modelId}...`);
            try {
              const response = await fetch(url, {
                method: 'POST',
                headers: {
                  'Authorization': `Bearer ${accessToken}`,
                  'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                  instances: [
                    {
                      messages: [
                        {
                          role: 'user',
                          content: fullPrompt
                        }
                      ]
                    }
                  ],
                  parameters: {
                    max_tokens: 4096,
                    temperature: temperature ? parseFloat(temperature) : 0.4
                  }
                })
              });

              if (response.ok) {
                data = await response.json();
                console.log(`Vertex AI prediction succeeded on region ${region} with model ${modelId}!`);
                break;
              } else {
                const errText = await response.text();
                console.log(`Region ${region} / Model ${modelId} failed: ${response.status} ${response.statusText} - ${errText}`);
                lastError = new Error(`Vertex AI Claude prediction error (${modelId}): ${response.status} ${response.statusText} - ${errText}`);
              }
            } catch (e: any) {
              console.log(`Network or request error for region ${region} / model ${modelId}:`, e.message);
              lastError = e;
            }
          }
          if (data) break;
        }
        
        if (!data) {
          throw lastError || new Error('All Vertex AI Claude candidate models and regions failed.');
        }

        console.log('Vertex AI Claude Meta-Evaluation complete.');

        // Safely extract prediction text
        let resultText = '';
        if (data.predictions && data.predictions[0]) {
          const pred = data.predictions[0];
          if (typeof pred.content === 'string') {
            resultText = pred.content;
          } else if (Array.isArray(pred.content)) {
            resultText = pred.content.map((c: any) => typeof c === 'string' ? c : c.text || '').join('');
          } else if (pred.content && typeof pred.content === 'object') {
            resultText = pred.content.text || JSON.stringify(pred.content);
          } else {
            resultText = JSON.stringify(pred);
          }
        } else {
          throw new Error('Invalid or empty predictions returned by Vertex AI.');
        }

        return NextResponse.json({
          success: true,
          result: resultText,
          usage: data.metadata?.usage || undefined
        });

      } catch (gcpError: any) {
        console.error('Vertex AI authentication or prediction failed:', gcpError);
        return NextResponse.json({ 
          success: false, 
          error: `Vertex AI Authentication/Access Error: ${gcpError.message}. \n\nTo use Claude via Google Cloud Vertex AI, you must configure Application Default Credentials (ADC) by running 'gcloud auth application-default login' in your local terminal. Alternatively, please enter a direct Anthropic API Key (sk-...) in the Scrutio settings sidebar.`
        }, { status: 400 });
      }
    }

    // Default Gemini path
    const isVertex = apiKey.startsWith('AQ.');
    const client = new GoogleGenAI(isVertex ? { apiKey, vertexai: true } : { apiKey });

    console.log(`Starting Level 2 Meta-Analysis using model ${model || 'gemini-3-flash-preview'}...`);

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
