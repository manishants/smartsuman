
'use server';

import {
  PdfToWordInputSchema,
  PdfToWordOutputSchema,
  type PdfToWordInput,
  type PdfToWordOutput,
  WordContentSchema,
} from '@/lib/types';
import { pdfToWordNoOcr } from '@/lib/actions/pdf-to-word-no-ocr';
import { ai } from '@/ai/genkit';
import { z } from 'zod';
import { Packer, Document, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell } from 'docx';

const ContentItemSchema = z.object({
  text: z.string().describe('The text content.'),
  bold: z.boolean().optional().describe('Whether the text is bold.'),
  italic: z.boolean().optional().describe('Whether the text is italic.'),
  color: z
    .string()
    .optional()
    .describe('The hex color of the text (e.g., #000000).'),
  fontSize: z.number().optional().describe('The font size of the text.'),
});

// Define a prompt to ensure consistent structured output from the model
const pdfToWordPrompt = ai.definePrompt({
  name: 'pdfToWordPrompt',
  model: 'googleai/gemini-2.0-flash',
  input: { schema: z.object({ pdfUri: z.string() }) },
  output: {
    schema: z.object({
      content: z.array(ContentItemSchema),
      structure: z.array(z.object({
        type: z.enum(['heading', 'paragraph', 'list', 'table', 'image', 'section_break', 'page_break']),
        level: z.number().optional().describe('Heading level (1-6) for headings'),
        items: z.array(z.string()).optional().describe('List items for bullet/numbered lists'),
        rows: z.array(z.array(z.string())).optional().describe('Table rows and cells'),
        imageDescription: z.string().optional().describe('Description of image content'),
        alignment: z.enum(['left', 'center', 'right', 'justify']).optional().describe('Text alignment'),
        position: z.object({
          x: z.number().optional().describe('X position on page'),
          y: z.number().optional().describe('Y position on page'),
          page: z.number().optional().describe('Page number'),
        }).optional(),
      })),
    }),
  },
  prompt: `Analyze the following PDF document COMPLETELY. Your task is to extract ALL content while preserving EXACT structure, layout, and styling. 

For TEXT content, provide:
1. The exact text content
2. Bold styling (true/false)
3. Italic styling (true/false)  
4. Font size in points
5. Color in hex format (e.g., #000000)
6. Text alignment (left/center/right/justify)
7. Position on page (x, y coordinates and page number)

For DOCUMENT STRUCTURE, identify and extract:
1. Headings (with level 1-6)
2. Paragraphs 
3. Lists (bullet/numbered with all items)
4. Tables (with all rows and cells)
5. Images (provide detailed descriptions)
6. Page breaks and section breaks
7. Headers and footers
8. Any visual elements, charts, or diagrams

Return BOTH the structured text content AND the complete document structure. Preserve the EXACT order and layout as it appears in the original PDF.

IMPORTANT: Extract EVERYTHING - do not miss any text, images, or design elements. The goal is perfect 1:1 conversion.

Document:
{{media url=pdfUri}}
`,
});

// Fallback prompt using Gemini 1.5 flash for wider availability
const pdfToWordPrompt15 = ai.definePrompt({
  name: 'pdfToWordPrompt15',
  model: 'googleai/gemini-1.5-flash',
  input: { schema: z.object({ pdfUri: z.string() }) },
  output: {
    schema: z.object({
      content: z.array(ContentItemSchema),
      structure: z.array(z.object({
        type: z.enum(['heading', 'paragraph', 'list', 'table', 'image', 'section_break', 'page_break']),
        level: z.number().optional().describe('Heading level (1-6) for headings'),
        items: z.array(z.string()).optional().describe('List items for bullet/numbered lists'),
        rows: z.array(z.array(z.string())).optional().describe('Table rows and cells'),
        imageDescription: z.string().optional().describe('Description of image content'),
        alignment: z.enum(['left', 'center', 'right', 'justify']).optional().describe('Text alignment'),
        position: z.object({
          x: z.number().optional().describe('X position on page'),
          y: z.number().optional().describe('Y position on page'),
          page: z.number().optional().describe('Page number'),
        }).optional(),
      })),
    }),
  },
  prompt: `Analyze the following PDF document COMPLETELY. Your task is to extract ALL content while preserving EXACT structure, layout, and styling. 

For TEXT content, provide:
1. The exact text content
2. Bold styling (true/false)
3. Italic styling (true/false)  
4. Font size in points
5. Color in hex format (e.g., #000000)
6. Text alignment (left/center/right/justify)
7. Position on page (x, y coordinates and page number)

For DOCUMENT STRUCTURE, identify and extract:
1. Headings (with level 1-6)
2. Paragraphs 
3. Lists (bullet/numbered with all items)
4. Tables (with all rows and cells)
5. Images (provide detailed descriptions)
6. Page breaks and section breaks
7. Headers and footers
8. Any visual elements, charts, or diagrams

Return BOTH the structured text content AND the complete document structure. Preserve the EXACT order and layout as it appears in the original PDF.

IMPORTANT: Extract EVERYTHING - do not miss any text, images, or design elements. The goal is perfect 1:1 conversion.

Document:
{{media url=pdfUri}}
`,
});

const pdfToWordFlow = ai.defineFlow(
  {
    name: 'pdfToWordFlow',
    inputSchema: PdfToWordInputSchema,
    outputSchema: z.object({
      content: z.array(ContentItemSchema),
      structure: z.array(z.object({
        type: z.enum(['heading', 'paragraph', 'list', 'table', 'image', 'section_break', 'page_break']),
        level: z.number().optional().describe('Heading level (1-6) for headings'),
        items: z.array(z.string()).optional().describe('List items for bullet/numbered lists'),
        rows: z.array(z.array(z.string())).optional().describe('Table rows and cells'),
        imageDescription: z.string().optional().describe('Description of image content'),
        alignment: z.enum(['left', 'center', 'right', 'justify']).optional().describe('Text alignment'),
        position: z.object({
          x: z.number().optional().describe('X position on page'),
          y: z.number().optional().describe('Y position on page'),
          page: z.number().optional().describe('Page number'),
        }).optional(),
      })),
    }),
  },
  async (input) => {
    const { pdfUri } = input;
    try {
      const { output } = await pdfToWordPrompt({ pdfUri });
      if (!output) {
        throw new Error('The AI failed to process the PDF content.');
      }
      return output;
    } catch (err) {
      // Try a fallback model for broader compatibility
      try {
        const { output } = await pdfToWordPrompt15({ pdfUri });
        if (!output) {
          throw new Error('The AI failed to process the PDF content (fallback).');
        }
        return output;
      } catch (err2) {
        const m1 = err instanceof Error ? err.message : String(err);
        const m2 = err2 instanceof Error ? err2.message : String(err2);
        throw new Error(
          `Gemini request failed: ${m1}. Fallback model also failed: ${m2}. ` +
            'Verify your API key is valid and enabled, and that outbound network access to generativelanguage.googleapis.com is allowed.'
        );
      }
    }
  }
);

export async function pdfToWord(
  input: PdfToWordInput
): Promise<PdfToWordOutput> {
  // If the user selected No OCR, use the local LibreOffice converter (no AI)
  if (input.conversionMode === 'no_ocr') {
    return pdfToWordNoOcr(input);
  }

  let result;
  try {
    result = await pdfToWordFlow({ pdfUri: input.pdfUri });
  } catch (e: any) {
    try {
      const fallback = await pdfToWordNoOcr(input);
      return fallback;
    } catch (e2: any) {
      throw new Error(
        e?.message || 'AI conversion failed and fallback converter was not available.'
      );
    }
  }

  // Create paragraphs from the structured content
  const paragraphs = result.content.map(
    (item) =>
      new Paragraph({
        children: [
          new TextRun({
            text: item.text,
            bold: item.bold,
            italics: item.italic,
            color: item.color?.substring(1), // Remove '#' from hex
            size: item.fontSize ? item.fontSize * 2 : 22, // Convert to half-points
          }),
        ],
      })
  );

  // Process document structure for enhanced formatting
  const structuredContent = result.structure?.map((struct) => {
    switch (struct.type) {
      case 'heading':
        return new Paragraph({
          heading: HeadingLevel[`HEADING_${struct.level || 1}` as keyof typeof HeadingLevel],
          children: [
            new TextRun({
              text: struct.items?.join(' ') || '',
              bold: true,
              size: 24 - (struct.level || 1) * 2, // Dynamic font size based on heading level
            }),
          ],
        });
      
      case 'list':
        return new Paragraph({
          children: struct.items?.map(item => 
            new TextRun({
              text: `• ${item}\n`,
            })
          ) || [],
        });
      
      case 'table':
        // Create proper table structure
        if (struct.rows && struct.rows.length > 0) {
          const tableRows = struct.rows.map(row => 
            new TableRow({
              children: row.map(cell => 
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: cell,
                        }),
                      ],
                    }),
                  ],
                })
              ),
            })
          );
          return new Table({
            rows: tableRows,
          });
        }
        return [];
      
      case 'image':
        // For images, we'll add a descriptive placeholder
        return new Paragraph({
          children: [
            new TextRun({
              text: `[Image: ${struct.imageDescription || 'Visual content'}]`,
              color: '0000FF',
              italics: true,
            }),
          ],
        });
      
      case 'page_break':
        return new Paragraph({
          children: [
            new TextRun({
              text: '\f', // Form feed for page break
            }),
          ],
        });
      
      default:
        return new Paragraph({
          children: [
            new TextRun({
              text: struct.items?.join(' ') || '',
            }),
          ],
        });
    }
  }).flat() || [];

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [...structuredContent, ...paragraphs],
      },
    ],
  });

  const docxBuffer = await Packer.toBuffer(doc);
  const docxUri = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${docxBuffer.toString(
    'base64'
  )}`;

  return { docxUri };
}
