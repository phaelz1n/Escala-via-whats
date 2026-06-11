const express = require('express');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.static(__dirname));
app.use('/temp_prints', express.static(path.join(__dirname, 'temp_prints')));

let browserContext = null;
let whatsappPage = null;
let isConnecting = false;

// Active Excel config persistence
const configPath = path.join(__dirname, 'config.json');
let activeExcelName = 'Escala 10.06.26.xlsx';

if (fs.existsSync(configPath)) {
  try {
    const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    if (config.activeExcelName) {
      activeExcelName = config.activeExcelName;
    }
  } catch (e) {
    console.error('Erro ao ler config.json:', e);
  }
}

function saveConfig() {
  try {
    fs.writeFileSync(configPath, JSON.stringify({ activeExcelName }, null, 2), 'utf8');
  } catch (e) {
    console.error('Erro ao salvar config.json:', e);
  }
}

function isRedCell(cell) {
  if (!cell || !cell.fill) return false;
  const fill = cell.fill;
  if (fill.type === 'pattern' && fill.fgColor) {
    const argb = fill.fgColor.argb;
    if (argb) {
      const hex = argb.toUpperCase();
      if (hex === 'FFFF0000' || hex === 'FFC00000' || hex.includes('FF0000') || hex === 'FFFFC7CE') {
        return true;
      }
    }
  }
  return false;
}

function formatTime(val) {
  if (val instanceof Date) {
    const hours = String(val.getUTCHours()).padStart(2, '0');
    const minutes = String(val.getUTCMinutes()).padStart(2, '0');
    return `${hours}:${minutes}`;
  }
  return val ? String(val).trim() : '';
}

function getCellValue(cell) {
  if (!cell || cell.value === null || cell.value === undefined) return '';
  let val = cell.value;
  if (typeof val === 'object') {
    if (val instanceof Date) {
      return formatTime(val);
    }
    if (val.result !== undefined) {
      val = val.result;
    } else if (val.richText) {
      return val.richText.map(t => t.text).join('');
    } else if (val.text) {
      return val.text;
    }
  }
  if (val instanceof Date) {
    return formatTime(val);
  }
  return String(val).trim();
}

async function generateScaleImages(excelPath, outputDir) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);

  const dadosSheet = workbook.getWorksheet('Dados') || workbook.worksheets.find(s => s.name.trim().toLowerCase() === 'dados');
  if (!dadosSheet) {
    throw new Error('Planilha "Dados" não encontrada no arquivo Excel.');
  }

  const driversList = [];
  dadosSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // Header
    const nameCell = row.getCell(1);
    const phoneCell = row.getCell(2);
    const name = getCellValue(nameCell);
    const phone = getCellValue(phoneCell);
    
    if (name && !isRedCell(nameCell)) {
      driversList.push({
        Name: name,
        Phone: phone,
        HasScale: false,
        ImagePath: null
      });
    }
  });

  const escalaSheet = workbook.worksheets.find(s => s.name.trim().toLowerCase().includes('escala di'));
  if (!escalaSheet) {
    throw new Error('Planilha de "Escala" não encontrada no arquivo Excel.');
  }

  const titleCell = escalaSheet.getRow(1).getCell(5);
  const titleText = getCellValue(titleCell) || 'ESCALA DE SERVIÇO';

  const headerRow = escalaSheet.getRow(2);
  const headers = [];
  for (let c = 1; c <= 6; c++) {
    headers.push(getCellValue(headerRow.getCell(c)));
  }

  const scaleRows = [];
  escalaSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber <= 2) return; // Skip headers
    const rowData = [];
    for (let c = 1; c <= 6; c++) {
      rowData.push(getCellValue(row.getCell(c)));
    }
    const driverName = rowData[5];
    if (driverName) {
      scaleRows.push({
        data: rowData,
        driverName: driverName.trim().toLowerCase()
      });
    }
  });

  if (fs.existsSync(outputDir)) {
    const files = fs.readdirSync(outputDir);
    for (const file of files) {
      try {
        fs.unlinkSync(path.join(outputDir, file));
      } catch (e) {
        console.error('Erro ao limpar imagem antiga:', e);
      }
    }
  } else {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const browser = await chromium.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });
  const page = await browser.newPage();

  for (const driver of driversList) {
    const matchingRows = scaleRows.filter(r => r.driverName === driver.Name.toLowerCase());
    if (matchingRows.length > 0) {
      driver.HasScale = true;

      let tableRowsHtml = '';
      matchingRows.forEach((r, idx) => {
        tableRowsHtml += `<tr class="${idx % 2 === 0 ? 'even' : 'odd'}">`;
        r.data.forEach((val, cIdx) => {
          let cellClass = '';
          if (cIdx === 3) cellClass = 'class="time-cell"';
          if (cIdx === 5) cellClass = 'class="driver-cell"';
          tableRowsHtml += `<td ${cellClass}>${val || ''}</td>`;
        });
        tableRowsHtml += '</tr>';
      });

      const htmlContent = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
          body {
            font-family: 'Outfit', sans-serif;
            margin: 0;
            padding: 0;
            background-color: transparent;
            display: flex;
            justify-content: center;
            align-items: center;
          }
          .scale-card {
            background: #0f172a;
            border: 2px dashed #38bdf8;
            border-radius: 16px;
            padding: 24px;
            width: 850px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5);
          }
          .header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            border-bottom: 2px solid rgba(56, 189, 248, 0.2);
            padding-bottom: 12px;
            margin-bottom: 18px;
          }
          .brand {
            color: #e2e8f0;
            font-weight: 700;
            font-size: 20px;
            letter-spacing: 1px;
          }
          .title {
            color: #38bdf8;
            font-weight: 600;
            font-size: 16px;
            text-transform: uppercase;
            background: rgba(56, 189, 248, 0.1);
            padding: 6px 12px;
            border-radius: 8px;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            color: #cbd5e1;
          }
          th {
            background-color: rgba(56, 189, 248, 0.15);
            color: #38bdf8;
            font-weight: 600;
            font-size: 12px;
            text-transform: uppercase;
            padding: 10px 14px;
            text-align: left;
            border: 1px solid rgba(255, 255, 255, 0.1);
          }
          td {
            padding: 10px 14px;
            font-size: 13px;
            border: 1px solid rgba(255, 255, 255, 0.08);
          }
          .even {
            background-color: rgba(255, 255, 255, 0.02);
          }
          .odd {
            background-color: transparent;
          }
          .time-cell {
            color: #f59e0b;
            font-weight: 700;
            font-family: monospace;
            font-size: 14px;
          }
          .driver-cell {
            color: #10b981;
            font-weight: 600;
          }
          .footer {
            margin-top: 20px;
            text-align: center;
            font-size: 11px;
            color: #64748b;
          }
        </style>
      </head>
      <body>
        <div class="scale-card" id="capture-target">
          <div class="header">
            <div class="brand"><span style="color: #38bdf8;">TRANS</span> PINHO</div>
            <div class="title">${titleText}</div>
          </div>
          <table>
            <thead>
              <tr>
                ${headers.map(h => `<th>${h || ''}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${tableRowsHtml}
            </tbody>
          </table>
          <div class="footer">Gerado automaticamente em ${new Date().toLocaleDateString('pt-BR')} - Trans Pinho</div>
        </div>
      </body>
      </html>
      `;

      await page.setContent(htmlContent);
      await page.waitForTimeout(100);
      const element = await page.$('#capture-target');
      const sanitizedName = driver.Name.replace(/[^a-zA-Z0-9_]/g, '_');
      const savePath = path.join(outputDir, `${sanitizedName}.png`);
      const webPath = `/temp_prints/${sanitizedName}.png`;

      await element.screenshot({ path: savePath, omitBackground: true });
      driver.ImagePath = savePath;
      driver.ImageUrl = webPath;
    }
  }

  await browser.close();
  return driversList;
}

// Check current connection status of WhatsApp
async function checkWhatsAppStatus() {
  if (!browserContext || !whatsappPage) return 'disconnected';
  try {
    const searchLocator = whatsappPage.locator('[data-testid="search"]');
    const qrLocator = whatsappPage.locator('canvas');
    
    const isSearchVisible = await searchLocator.isVisible();
    const isQrVisible = await qrLocator.isVisible();
    
    if (isSearchVisible) {
      return 'connected';
    } else if (isQrVisible) {
      return 'qr_ready';
    } else {
      return 'loading';
    }
  } catch (e) {
    return 'disconnected';
  }
}

// Endpoint to get WhatsApp status
app.get('/api/whatsapp/status', async (req, res) => {
  const status = await checkWhatsAppStatus();
  res.json({ status });
});

// Endpoint to start WhatsApp headless browser
app.post('/api/whatsapp/start', async (req, res) => {
  if (isConnecting) {
    return res.status(400).json({ error: 'Já existe uma tentativa de conexão em andamento.' });
  }
  
  const status = await checkWhatsAppStatus();
  if (status === 'connected' || status === 'qr_ready' || status === 'loading') {
    return res.json({ success: true, status });
  }
  
  isConnecting = true;
  try {
    const userDataDir = path.join(__dirname, 'whatsapp_session');
    
    browserContext = await chromium.launchPersistentContext(userDataDir, {
      headless: true,
      viewport: { width: 1280, height: 800 },
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--disable-gpu'
      ]
    });
    
    const pages = browserContext.pages();
    whatsappPage = pages.length > 0 ? pages[0] : await browserContext.newPage();
    
    await whatsappPage.goto('https://web.whatsapp.com');
    
    res.json({ success: true, message: 'Navegador WhatsApp Web iniciado em segundo plano.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Falha ao iniciar WhatsApp: ' + e.message });
  } finally {
    isConnecting = false;
  }
});

// Endpoint to get WhatsApp QR Code as Base64 Image
app.get('/api/whatsapp/qr', async (req, res) => {
  if (!whatsappPage) {
    return res.status(400).json({ error: 'WhatsApp não está iniciado.' });
  }
  try {
    const status = await checkWhatsAppStatus();
    if (status === 'qr_ready') {
      const qrLocator = whatsappPage.locator('canvas');
      if (await qrLocator.isVisible()) {
        const qrBuffer = await qrLocator.screenshot({ type: 'png' });
        const base64 = qrBuffer.toString('base64');
        return res.json({ qr: `data:image/png;base64,${base64}`, status });
      }
    }
    res.json({ status });
  } catch (e) {
    res.status(500).json({ error: 'Erro ao capturar QR Code: ' + e.message });
  }
});

// Endpoint to stop WhatsApp browser
app.post('/api/whatsapp/stop', async (req, res) => {
  try {
    if (browserContext) {
      await browserContext.close();
      browserContext = null;
      whatsappPage = null;
      return res.json({ success: true, message: 'Navegador fechado.' });
    }
    res.json({ success: true, message: 'O navegador já estava fechado.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Erro ao fechar navegador: ' + e.message });
  }
});

// Endpoint to get active excel name
app.get('/api/active-excel', (req, res) => {
  res.json({ filename: activeExcelName });
});

const gitPath = fs.existsSync('C:\\Program Files\\Git\\cmd\\git.exe') ? 'C:\\Program Files\\Git\\cmd\\git.exe' : 'git';

function gitAutoCommit(filename) {
  const gitCmdAdd = `"${gitPath}" add -A`;
  const commitMessage = `Upload e atualizacao de escala: ${filename}`;
  const gitCmdCommit = `"${gitPath}" commit -m "${commitMessage}"`;
  
  exec(gitCmdAdd, { cwd: __dirname }, (addErr) => {
    if (addErr) {
      console.error('Erro ao dar git add:', addErr);
      return;
    }
    exec(gitCmdCommit, { cwd: __dirname }, (commitErr, stdout) => {
      if (commitErr) {
        const msg = commitErr.message || '';
        const isClean = msg.includes('nothing to commit') || msg.includes('no changes added to commit') || msg.includes('clean');
        if (!isClean) {
          console.error('Erro ao dar git commit:', commitErr);
        }
      } else {
        console.log('Git auto-committed com sucesso:', stdout.trim());
      }
    });
  });
}

// Endpoint to upload a new excel sheet via Base64 JSON
app.post('/api/upload', (req, res) => {
  const { filename, base64 } = req.body;
  if (!filename || !base64) {
    return res.status(400).json({ error: 'Parâmetros inválidos (filename e base64 são necessários).' });
  }

  try {
    const filePath = path.join(__dirname, filename);
    const buffer = Buffer.from(base64, 'base64');
    
    fs.writeFileSync(filePath, buffer);
    
    activeExcelName = filename;
    saveConfig();
    
    // Auto-commit the changes to Git in the background
    gitAutoCommit(filename);
    
    res.json({ success: true, filename: activeExcelName, message: 'Planilha atualizada com sucesso.' });
  } catch (e) {
    console.error('Upload error:', e);
    res.status(500).json({ error: 'Erro ao salvar o arquivo Excel: ' + e.message });
  }
});

// Endpoint to parse drivers and generate schedule prints
app.get('/api/drivers', async (req, res) => {
  const excelPath = path.join(__dirname, activeExcelName);
  
  if (!fs.existsSync(excelPath)) {
    return res.status(404).json({ error: `O arquivo Excel ativo não foi encontrado: ${activeExcelName}. Por favor, envie uma nova planilha.` });
  }
  
  const outputDir = path.join(__dirname, 'temp_prints');
  
  try {
    const drivers = await generateScaleImages(excelPath, outputDir);
    res.json({ success: true, drivers });
  } catch (e) {
    console.error('Erro ao processar planilha de escalas:', e);
    res.status(500).json({ error: 'Falha ao processar planilha de escalas: ' + e.message });
  }
});

// Endpoint to send WhatsApp message to a driver
app.post('/api/whatsapp/send', async (req, res) => {
  const { name, phone, text, imagePath } = req.body;
  
  if (!phone || !text || !imagePath) {
    return res.status(400).json({ error: 'Dados insuficientes para o envio (telefone, texto e imagem são obrigatórios).' });
  }
  
  const status = await checkWhatsAppStatus();
  if (status !== 'connected') {
    return res.status(400).json({ error: 'O WhatsApp não está conectado. Conecte-o e escaneie o QR Code primeiro.' });
  }
  
  try {
    // Format phone number
    let cleanPhone = phone.replace(/\D/g, '');
    if (cleanPhone.length <= 11) {
      cleanPhone = '5551' + cleanPhone; // Fallback default country code 55 + DDD 51
    } else if (cleanPhone.length === 12 && cleanPhone.startsWith('51')) {
      cleanPhone = '55' + cleanPhone; // Add country code if only DDD starts
    }
    
    const absoluteImagePath = path.resolve(imagePath);
    if (!fs.existsSync(absoluteImagePath)) {
      return res.status(404).json({ error: 'Arquivo de imagem não encontrado no servidor: ' + imagePath });
    }
    
    // Navigate to contact send link
    await whatsappPage.goto(`https://web.whatsapp.com/send?phone=${cleanPhone}`);
    
    const chatInputSelector = 'div[contenteditable="true"]';
    const invalidDialogSelector = 'div[role="button"]:has-text("OK"), button:has-text("OK"), [data-testid="popup-controls-ok"]';
    
    let chatLoaded = false;
    let chatError = null;
    
    // Wait for chat or invalid number modal
    for (let i = 0; i < 25; i++) {
      const isInput = await whatsappPage.locator(chatInputSelector).first().isVisible();
      const isInvalid = await whatsappPage.locator(invalidDialogSelector).first().isVisible();
      
      if (isInput) {
        chatLoaded = true;
        break;
      }
      if (isInvalid) {
        chatError = `O número ${cleanPhone} não está cadastrado no WhatsApp.`;
        await whatsappPage.locator(invalidDialogSelector).first().click().catch(() => {});
        break;
      }
      await whatsappPage.waitForTimeout(1000);
    }
    
    if (!chatLoaded) {
      throw new Error(chatError || 'Tempo limite excedido aguardando o chat carregar.');
    }
    
    await whatsappPage.waitForTimeout(1000);
    
    const plusBtnSelector = 'span[data-icon="plus"], span[data-icon="attach-menu-plus"], div[aria-label="Anexar"]';
    const fileInputSelector = 'input[type="file"]';
    
    let fileInputs = await whatsappPage.locator(fileInputSelector).all();
    if (fileInputs.length === 0) {
      await whatsappPage.locator(plusBtnSelector).first().click().catch(() => {});
      await whatsappPage.waitForTimeout(1000);
      fileInputs = await whatsappPage.locator(fileInputSelector).all();
    }
    
    if (fileInputs.length === 0) {
      throw new Error('Não foi possível localizar o botão de anexo.');
    }
    
    // Use the file input that accepts images
    let targetInput = fileInputs[0];
    for (const input of fileInputs) {
      const acceptAttr = await input.getAttribute('accept').catch(() => '');
      if (acceptAttr && acceptAttr.includes('image')) {
        targetInput = input;
        break;
      }
    }
    
    await targetInput.setInputFiles(absoluteImagePath);
    
    // Wait for the media send preview screen
    const sendMediaBtnSelector = 'span[data-icon="send"], [data-testid="send"]';
    await whatsappPage.waitForSelector(sendMediaBtnSelector, { timeout: 10000 });
    
    // Insert caption
    const captionInput = whatsappPage.locator('div[contenteditable="true"]').last();
    if (await captionInput.isVisible()) {
      await captionInput.focus();
      await whatsappPage.keyboard.type(text);
      await whatsappPage.waitForTimeout(500);
    }
    
    // Click send
    await whatsappPage.locator(sendMediaBtnSelector).first().click();
    
    // Wait for message sending completion
    await whatsappPage.waitForTimeout(4000);
    
    res.json({ success: true, message: `Mensagem enviada com sucesso para ${name}.` });
  } catch (e) {
    console.error(`Erro ao enviar para ${name}:`, e);
    res.status(500).json({ error: `Falha ao enviar para ${name}: ` + e.message });
  }
});

// Start express server
const server = app.listen(PORT, () => {
  console.log(`Servidor local rodando em http://localhost:${PORT}`);
});
server.timeout = 600000; // 10 minutes timeout to handle large excel files
