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
    const requests = buildUSFCRFormattingRequests(structure);

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

function buildUSFCRFormattingRequests(structure) {
  const requests = [];
  let currentIndex = 1;

  // USFCR BRAND TYPOGRAPHY SYSTEM
  // Dark blue for headers: rgb(26, 35, 50) = #1A2332
  const usfcrDarkBlue = { red: 0.102, green: 0.137, blue: 0.196 };
  const bodyGray = { red: 0.15, green: 0.15, blue: 0.15 };

  const styles = {
    heading1: {
      fontSize: 18,
      bold: true,
      fontFamily: 'Montserrat',
      color: usfcrDarkBlue,
      lineSpacing: 120,
      spaceAbove: 0,
      spaceBelow: 16
    },
    heading2: {
      fontSize: 14,
      bold: true,
      fontFamily: 'Montserrat',
      color: usfcrDarkBlue,
      lineSpacing: 130,
      spaceAbove: 18,
      spaceBelow: 10
    },
    heading3: {
      fontSize: 12,
      bold: true,
      fontFamily: 'Montserrat',
      color: usfcrDarkBlue,
      lineSpacing: 130,
      spaceAbove: 14,
      spaceBelow: 8
    },
    body: {
      fontSize: 11,
      bold: false,
      fontFamily: 'Avenir',
      color: bodyGray,
      lineSpacing: 160,
      spaceAbove: 0,
      spaceBelow: 12
    }
  };

  requests.push({
    updateDocumentStyle: {
      documentStyle: {
        marginTop: { magnitude: 72, unit: 'PT' },
        marginBottom: { magnitude: 72, unit: 'PT' },
        marginLeft: { magnitude: 90, unit: 'PT' },
        marginRight: { magnitude: 90, unit: 'PT' },
        defaultHeaderId: '',
        defaultFooterId: ''
      },
      fields: 'marginTop,marginBottom,marginLeft,marginRight'
    }
  });

  structure.forEach((element) => {
    const text = element.text + '\n';
    const textLength = text.length;

    requests.push({
      insertText: {
        text: text,
        location: { index: currentIndex }
      }
    });

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
          weightedFontFamily: {
            fontFamily: style.fontFamily,
            weight: style.bold ? 600 : 400
          },
          foregroundColor: {
            color: {
              rgbColor: style.color
            }
          }
        },
        range: {
          startIndex: currentIndex,
          endIndex: currentIndex + textLength - 1
        },
        fields: 'fontSize,bold,weightedFontFamily,foregroundColor'
      }
    });

    requests.push({
      updateParagraphStyle: {
        paragraphStyle: {
          spaceAbove: { magnitude: style.spaceAbove, unit: 'PT' },
          spaceBelow: { magnitude: style.spaceBelow, unit: 'PT' },
          lineSpacing: style.lineSpacing,
          alignment: 'START',
          avoidWidowAndOrphan: true,
          keepLinesTogether: element.type === 'heading',
          keepWithNext: element.type === 'heading'
        },
        range: {
          startIndex: currentIndex,
          endIndex: currentIndex + textLength
        },
        fields: 'spaceAbove,spaceBelow,lineSpacing,alignment,avoidWidowAndOrphan,keepLinesTogether,keepWithNext'
      }
    });

    currentIndex += textLength;
  });

  return requests;
}