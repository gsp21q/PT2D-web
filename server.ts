import express from "express";
import { marked } from "marked";
import { createServer as createViteServer } from "vite";
import path from "path";
import { createRequire } from "module";
const customRequire = createRequire(import.meta.url);
const { Document, Packer, Paragraph, TextRun, AlignmentType, XmlComponent, XmlAttributeComponent } = customRequire("docx");
const { convertLatex2Math, mathJaxReady } = customRequire("@hungknguyen/docx-math-converter");

const PORT = 3000;

class RawXml extends XmlComponent {
  constructor(key: string, root: any[]) {
    super(key);
    (this as any).root = root;
  }
}

class RawAttr extends XmlAttributeComponent {
  constructor(xmlKeys: any, root: any) {
    super(root);
    (this as any).xmlKeys = xmlKeys;
    (this as any).root = root;
  }
}

function deepPatchMath(obj: any): any {
  if (Array.isArray(obj)) {
    return obj.map(deepPatchMath);
  } else if (obj && typeof obj === 'object') {
    if (obj.rootKey === '_attr') {
      return new RawAttr(obj.xmlKeys, obj.root);
    } else if (obj.rootKey) {
      const children = obj.root ? deepPatchMath(obj.root) : [];
      return new RawXml(obj.rootKey, children);
    }
  }
  return obj;
}

// Custom extension for inline math ($...$)
const inlineMathExtension = {
  name: 'inlineMath',
  level: 'inline',
  start(src: string) { return src.match(/\$/)?.index; },
  tokenizer(src: string, tokens: any) {
    const match = /^\$([^$\n]+)\$/.exec(src);
    if (match) {
      return {
        type: 'inlineMath',
        raw: match[0],
        math: match[1],
      };
    }
  },
  renderer(token: any) { return token.raw; }
};

// Custom extension for block math ($$...$$)
const blockMathExtension = {
  name: 'blockMath',
  level: 'block',
  start(src: string) { return src.match(/\$\$/)?.index; },
  tokenizer(src: string, tokens: any) {
    const match = /^\$\$([^$]+)\$\$/.exec(src);
    if (match) {
      return {
        type: 'blockMath',
        raw: match[0],
        math: match[1],
      };
    }
  },
  renderer(token: any) { return token.raw; }
};

marked.use({ extensions: [inlineMathExtension, blockMathExtension] });

async function createDocxBuffer(text: string, direction: string, colors: any) {
  const isRtl = direction === 'rtl';

  function processTokens(tokensList: any[]) {
    const runs: any[] = [];
    for (const t of tokensList) {
      if (t.type === 'text' || t.type === 'escape' || t.type === 'html') {
        runs.push(new TextRun({ text: t.text || t.raw, rightToLeft: isRtl }));
      } else if (t.type === 'strong') {
        runs.push(new TextRun({ text: t.text, bold: true, rightToLeft: isRtl }));
      } else if (t.type === 'em') {
        runs.push(new TextRun({ text: t.text, italics: true, rightToLeft: isRtl }));
      } else if (t.type === 'inlineMath') {
        const mathObj = convertLatex2Math(t.math);
        runs.push(deepPatchMath(mathObj));
      } else if (t.type === 'codespan') {
        runs.push(new TextRun({ text: t.text, font: "Courier New", rightToLeft: isRtl }));
      }
    }
    return runs;
  }

  // Pre-process $$math$$ so marked does not wrap them in paragraphs inside another paragraph
  const processedText = text.replace(/\$\$(.*?)\$\$/gs, '\n\n$$$$$1$$$$\n\n');
  const tokens = marked.lexer(processedText);
  const children: any[] = [];

  for (const t of tokens) {
    if (t.type === 'heading') {
      const runs = processTokens(t.tokens || []);
      const styleId = t.depth === 1 ? 'Heading1' : t.depth === 2 ? 'Heading2' : 'Normal';
      children.push(new Paragraph({ style: styleId, children: runs, bidirectional: isRtl }));
    } else if (t.type === 'paragraph') {
      const runs = processTokens(t.tokens || []);
      children.push(new Paragraph({ style: 'BodyText', children: runs, bidirectional: isRtl }));
    } else if (t.type === 'blockMath') {
      const mathObj = convertLatex2Math(t.math);
      const patchedMath = deepPatchMath(mathObj);
      
      const displayMath = new RawXml('m:oMathPara', [
        new RawXml('m:oMathParaPr', [
          new RawXml('m:jc', [
            new RawAttr({ val: 'm:val' }, { val: 'center' })
          ])
        ]),
        patchedMath
      ]);

      children.push(new Paragraph({ 
        style: 'BodyText', 
        children: [displayMath],
        alignment: AlignmentType.CENTER 
      }));
    }
  }

  const paragraphStyles = [
    {
      id: "Heading1",
      name: "Heading 1",
      basedOn: "Normal",
      next: "Normal",
      run: { color: colors.titleColor.replace('#', '') }
    },
    {
      id: "Heading2",
      name: "Heading 2",
      basedOn: "Normal",
      next: "Normal",
      run: { color: colors.subtitleColor.replace('#', '') }
    },
    {
      id: "Normal",
      name: "Normal",
      quickFormat: true,
      run: { color: colors.bodyColor.replace('#', '') }
    },
    {
      id: "BodyText",
      name: "Body Text",
      basedOn: "Normal",
      next: "Normal",
      run: { color: colors.bodyColor.replace('#', '') }
    }
  ];

  const doc = new Document({
    creator: "PT2D",
    styles: { paragraphStyles },
    sections: [{ properties: {}, children }]
  });

  return await Packer.toBuffer(doc);
}

async function startServer() {
  await mathJaxReady(); // Wait for MathJax to initialize once
  const app = express();

  app.use(express.json({ limit: "50mb" }));

  // Middleware to add CORS for external host environments
  app.use((req, res, next) => {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
  });

  app.post("/api/convert", async (req, res) => {
    try {
      const { 
        text, 
        direction = "ltr", 
        titleColor = "#000000", 
        subtitleColor = "#000000", 
        bodyColor = "#000000" 
      } = req.body;

      console.log("Converting using docx + mathematical elements in memory...");
      
      // Generate the Document Buffer in memory!
      const buffer = await createDocxBuffer(text, direction, { titleColor, subtitleColor, bodyColor });

      // Return the buffer as a stream
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      res.setHeader("Content-Disposition", `attachment; filename="document.docx"`);
      res.send(buffer);
    } catch (e: any) {
      console.error(e);
      res.status(500).json({ error: e.message || "Failed to convert" });
    }
  });

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
