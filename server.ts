import express from 'express';
import path from 'path';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI } from '@google/genai';

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json({ limit: '50mb' }));

  // API Route for Gemini analysis
  app.post('/api/analyze', async (req, res) => {
    try {
      const { dataText } = req.body;
      if (!dataText) {
        return res.status(400).json({ error: 'No data provided' });
      }

      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey || apiKey === 'MY_GEMINI_API_KEY') {
        return res.status(401).json({ error: 'Please configure your Gemini API Key in the AI Studio Settings (Secrets panel).' });
      }

      const ai = new GoogleGenAI({ apiKey });

      const prompt = `Act as a CFO. Analyze the following Product Line data:

${dataText}

Nhiệm vụ:
1. Tuỳ theo dữ liệu hiện có (Doanh số, Lợi nhuận, vv), phân tích phân khúc hoặc đối tượng mang lại kết quả tốt nhất.
2. Nếu có dữ liệu chi phí giá vốn (COGS), chỉ ra khu vực/ngành hàng có COGS cao bất thường so với doanh thu.
3. Nếu có dữ liệu giới tính hoặc địa điểm, nhận xét về xu hướng mua hàng của các tệp khách hàng / khu vực đó.
4. Đề xuất ngắn gọn 2 chiến lược quản trị nội bộ hoặc phát triển kinh doanh.

Yêu cầu định dạng đầu ra: Không viết thành một báo cáo hoàn chỉnh dài dòng. Chỉ cần trình bày trực tiếp bằng các gạch đầu dòng các chỉ số KHẢ DỤNG đã phân tích và các khuyến nghị. Trình bày toàn bộ phần phân tích bằng Tiếng Việt trước, sau đó là bản dịch toàn bộ sang Tiếng Anh.`;

      let responseText = '';
      let retries = 5;
      let delay = 3000;
      let modelToUse = 'gemini-2.5-flash';
      while (retries > 0) {
        try {
          const response = await ai.models.generateContent({
            model: modelToUse,
            contents: prompt,
          });
          responseText = response.text || '';
          break;
        } catch (e: any) {
          retries--;
          
          if (retries === 0) {
             if (e?.status === 429 || String(e?.message).includes('Exceeded')) {
                 throw new Error('API Rate Limit Exceeded (429). Please configure your own Gemini API Key in the Settings (Secrets panel) or try again later.');
             }
             throw e;
          }
          if (e?.status === 429 || e?.status === 503 || String(e?.message).includes('demand')) {
             delay = Math.max(delay, 5000);
          }
          if (retries <= 3) modelToUse = 'gemini-1.5-flash'; 
          if (retries <= 1) modelToUse = 'gemini-2.0-flash';
          await new Promise(res => setTimeout(res, delay));
          delay *= 1.5;
        }
      }

      res.json({ result: responseText });
    } catch (error: any) {
      
      res.status(500).json({ error: error.message || 'An error occurred during analysis.' });
    }
  });

  // API Route for Follow-up Chat
  app.post('/api/chat', async (req, res) => {
    try {
      const { dataText, question, history } = req.body;
      if (!dataText || !question) {
        return res.status(400).json({ error: 'Missing dataText or question' });
      }

      const apiKey = process.env.GEMINI_API_KEY;
      if (!apiKey || apiKey === 'MY_GEMINI_API_KEY') {
        return res.status(401).json({ error: 'Please configure your Gemini API Key in the AI Studio Settings (Secrets panel).' });
      }

      const ai = new GoogleGenAI({ apiKey });

      const historyText = history ? history.map((h: any) => `${h.role === 'user' ? 'User' : 'CFO'}: ${h.content}`).join('\n') : '';

      const prompt = `Act as a CFO. You are having a conversation about the following financial data:

${dataText}

Conversation History:
${historyText}

User Question: ${question}

Respond professionally. Provide your full response in Vietnamese first, then provide the full English translation. Limit your overall response to be concise and focused specifically on the data and the question asked.`;

      let responseText = '';
      let retries = 5;
      let delay = 3000;
      let modelToUse = 'gemini-2.5-flash';
      while (retries > 0) {
        try {
          const response = await ai.models.generateContent({
            model: modelToUse,
            contents: prompt,
          });
          responseText = response.text || '';
          break;
        } catch (e: any) {
          retries--;
          
          if (retries === 0) {
             if (e?.status === 429 || String(e?.message).includes('Exceeded')) {
                 throw new Error('API Rate Limit Exceeded (429). Please configure your own Gemini API Key in the Settings (Secrets panel) or try again later.');
             }
             throw e;
          }
          if (e?.status === 429 || e?.status === 503 || String(e?.message).includes('demand')) {
             delay = Math.max(delay, 5000);
          }
          if (retries <= 3) modelToUse = 'gemini-1.5-flash';
          if (retries <= 1) modelToUse = 'gemini-2.0-flash';
          await new Promise(res => setTimeout(res, delay));
          delay *= 1.5;
        }
      }

      res.json({ result: responseText });
    } catch (error: any) {
      
      res.status(500).json({ error: error.message || 'An error occurred during chat.' });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
