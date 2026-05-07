import { Document, Packer, Paragraph, TextRun, AlignmentType, XmlComponent, XmlAttributeComponent, Table, TableRow, TableCell, WidthType, BorderStyle } from "docx";
import { convertLatex2Math, mathJaxReady } from "@hungknguyen/docx-math-converter";
import { marked } from "marked";

class RawXml extends XmlComponent {
  constructor(key: string, root: any[]) {
    super(key);
    (this as any).root = root;
  }
}

class RawAttr extends XmlAttributeComponent<any> {
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

export async function createDocxBuffer(text: string, direction: string, colors: any) {
  const isRtl = direction === 'rtl';

  function processTokens(tokensList: any[], options: { bold?: boolean; italic?: boolean } = {}) {
    const runs: any[] = [];
    for (const t of tokensList) {
      if (t.type === 'paragraph' || (t.type === 'text' && t.tokens)) {
        runs.push(...processTokens(t.tokens, options));
      } else if (t.type === 'text' || t.type === 'escape' || t.type === 'html') {
        runs.push(new TextRun({ text: t.text || t.raw || "", bold: options.bold, italics: options.italic, rightToLeft: isRtl }));
      } else if (t.type === 'strong') {
        runs.push(...processTokens(t.tokens || [], { ...options, bold: true }));
      } else if (t.type === 'em') {
        runs.push(...processTokens(t.tokens || [], { ...options, italic: true }));
      } else if (t.type === 'inlineMath' || t.type === 'blockMath') {
        const mathObj = convertLatex2Math(t.math);
        runs.push(deepPatchMath(mathObj));
      } else if (t.type === 'codespan') {
        runs.push(new TextRun({ text: t.text, font: "Courier New", bold: options.bold, italics: options.italic, rightToLeft: isRtl }));
      } else if (t.type === 'br') {
        runs.push(new TextRun({ break: 1 }));
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
    } else if (t.type === 'hr') {
      children.push(new Paragraph({
        children: [],
        border: {
          bottom: {
            color: 'auto',
            space: 1,
            style: BorderStyle.SINGLE,
            size: 6,
          },
        },
      }));
    } else if (t.type === 'list') {
      for (const item of t.items) {
        const runs = processTokens(item.tokens || []);
        children.push(new Paragraph({
          style: 'BodyText',
          children: runs,
          bidirectional: isRtl,
          bullet: t.ordered ? undefined : { level: 0 },
          numbering: t.ordered ? { reference: "default-ordered", level: 0 } : undefined
        }));
      }
    } else if (t.type === 'table') {
      const rows = [];
      
      if (t.header && t.header.length > 0) {
        const cells = t.header.map((th: any) => {
          const runs = processTokens(th.tokens || [], { bold: true });
          return new TableCell({
            children: [new Paragraph({ children: runs, bidirectional: isRtl })],
          });
        });
        rows.push(new TableRow({ children: cells, tableHeader: true }));
      }

      if (t.rows && t.rows.length > 0) {
        for (const tr of t.rows) {
          const cells = tr.map((td: any) => {
            const runs = processTokens(td.tokens || []);
            return new TableCell({
              children: [new Paragraph({ children: runs, bidirectional: isRtl })],
            });
          });
          rows.push(new TableRow({ children: cells }));
        }
      }

      children.push(new Table({
        rows,
        alignment: AlignmentType.CENTER,
        width: { size: 0, type: WidthType.AUTO },
      }));
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
    numbering: {
      config: [
        {
          reference: "default-ordered",
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: { left: 720, hanging: 360 },
                },
              },
            },
          ],
        },
      ],
    },
    sections: [{ properties: {}, children }]
  });

  return await Packer.toBlob(doc);
}
