import express from "express";
import { convert } from "pandoc-wasm";
import { Document, Packer, Paragraph } from "docx";
import { createServer as createViteServer } from "vite";
import path from "path";

const PORT = 3000;

async function createReferenceDocx(colors: { title?: string; subtitle?: string; body?: string }) {
  const paragraphStyles = [];

  if (colors.title) {
    paragraphStyles.push({
      id: "Heading1",
      name: "Heading 1",
      basedOn: "Normal",
      next: "Normal",
      run: { color: colors.title.replace('#', '') }
    });
  }
  if (colors.subtitle) {
    paragraphStyles.push({
      id: "Heading2",
      name: "Heading 2",
      basedOn: "Normal",
      next: "Normal",
      run: { color: colors.subtitle.replace('#', '') }
    });
  }
  if (colors.body) {
    paragraphStyles.push({
      id: "Normal",
      name: "Normal",
      quickFormat: true,
      run: { color: colors.body.replace('#', '') }
    });
    paragraphStyles.push({
      id: "BodyText",
      name: "Body Text",
      basedOn: "Normal",
      next: "Normal",
      run: { color: colors.body.replace('#', '') }
    });
  }

  // To make a valid docx, it needs at least one empty section
  const doc = new Document({
    creator: "Applet",
    styles: {
      paragraphStyles
    },
    sections: [
      {
        properties: {},
        children: [new Paragraph("")]
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);
  return new Blob([buffer]);
}

async function startServer() {
  const app = express();

  app.use(express.json({ limit: "50mb" }));

  app.post("/api/convert", async (req, res) => {
    try {
      const { 
        text, 
        direction = "ltr", 
        titleColor = "", 
        subtitleColor = "", 
        bodyColor = "" 
      } = req.body;

      // Ensure $$ and $ are processed as math
      // Pandoc natively handles this with markdown+tex_math_dollars
      
      let markdown = `---\n`;
      markdown += `dir: ${direction}\n`;
      markdown += `---\n\n`;
      markdown += text;

      const options: any = {
        from: "markdown+tex_math_dollars",
        to: "docx",
        "output-file": "output.docx",
        standalone: true
      };

      const files: any = {};

      if (titleColor || subtitleColor || bodyColor) {
        const refBlob = await createReferenceDocx({
          title: titleColor,
          subtitle: subtitleColor,
          body: bodyColor
        });
        files["custom-reference.docx"] = refBlob;
        options["reference-doc"] = "custom-reference.docx";
      }

      console.log("Converting using pandoc-wasm...");
      const result = await convert(options, markdown, files);
      
      const outBlob = result.files["output.docx"];
      if (!outBlob) {
        throw new Error("Conversion failed to generate output.docx");
      }

      const buffer = Buffer.from(await outBlob.arrayBuffer());

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
