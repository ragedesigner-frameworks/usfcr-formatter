export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader('Access-Control-Allow-Headers', 'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version, Authorization');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    const { accessToken, structure, title } = req.body;

    if (!accessToken || !structure) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const createResponse = await fetch('https://docs.googleapis.com/v1/documents', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ title: title || 'Formatted Document' })
    });

    if (!createResponse.ok) {
      const error = await createResponse.json();
      return res.status(createResponse.status).json({ error: 'Failed to create document', details: error });
    }

    const doc = await createResponse.json();
    const documentId = doc.documentId;
    const requests = buildFormattingRequests(structure);

    const updateResponse = await fetch(`https://docs.googleapis.com/v1/documents/${documentId}:batchUpdate`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ requests })
    });

    if (!updateResponse.ok) {
      const error = await updateResponse.json();
      return res.status(updateResponse.status).json({ error: 'Failed to format document', details: error });
    }

    return res.status(200).json({
      success: true,
      documentId: documentId,
      url: `https://docs.google.com/document/d/${documentId}/edit`
    });

  } catch (error) {
    console.error('Document creation error:', error);
    return res.status(500).json({ error: 'Internal server error', message: error.message });
  }
}

function buildFormattingRequests(structure) {
  const requests = [];
  let currentIndex = 1;
  const styles = {
    heading1: { fontSize: 14, bold: true, color: { red: 0.1, green: 0.16, blue: 0.2 } },
    heading2: { fontSize: 12, bold: true, color: { red: 0.2, green: 0.2, blue: 0.2 } },
    heading3: { fontSize: 11, bold: true, color: { red: 0.3, green: 0.3, blue: 0.3 } },
    body: { fontSize: 11, bold: false, color: { red: 0.18, green: 0.18, blue: 0.18 } }
  };
  structure.forEach((element) => {
    const text = element.text + '\n';
    const textLength = text.length;
    requests.push({ insertText: { text: text, location: { index: currentIndex } } });
    let style;
    if (element.type === 'heading') {
      const styleKey = element.level === 1 ? 'heading1' : element.level === 2 ? 'heading2' : 'heading3';
      style = styles[styleKey];
    } else {
      style = styles.body;
    }
    requests.push({
      updateTextStyle: {
        textStyle: {
          fontSize: { magnitude: style.fontSize, unit: 'PT' },
          bold: style.bold,
          foregroundColor: { color: { rgbColor: style.color } }
        },
        range: { startIndex: currentIndex, endIndex: currentIndex + textLength - 1 },
        fields: 'fontSize,bold,foregroundColor'
      }
    });
    const spaceAbove = element.type === 'heading' ? 12 : 0;
    const spaceBelow = element.type === 'heading' ? 6 : 0;
    requests.push({
      updateParagraphStyle: {
        paragraphStyle: {
          spaceAbove: { magnitude: spaceAbove, unit: 'PT' },
          spaceBelow: { magnitude: spaceBelow, unit: 'PT' },
          lineSpacing: 140
        },
        range: { startIndex: currentIndex, endIndex: currentIndex + textLength },
        fields: 'spaceAbove,spaceBelow,lineSpacing'
      }
    });
    currentIndex += textLength;
  });
  return requests;
}