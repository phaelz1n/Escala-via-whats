const express = require('express');
const { exec } = require('child_process');
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright');

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

// Endpoint to start WhatsApp headed browser
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
      headless: false,
      executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
      viewport: null,
      args: ['--start-maximized']
    });
    
    const pages = browserContext.pages();
    whatsappPage = pages.length > 0 ? pages[0] : await browserContext.newPage();
    
    await whatsappPage.goto('https://web.whatsapp.com');
    
    res.json({ success: true, message: 'Navegador WhatsApp Web iniciado.' });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: 'Falha ao iniciar WhatsApp: ' + e.message });
  } finally {
    isConnecting = false;
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

const gitPath = 'C:\\Program Files\\Git\\cmd\\git.exe';

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
app.get('/api/drivers', (req, res) => {
  const excelPath = path.join(__dirname, activeExcelName);
  
  if (!fs.existsSync(excelPath)) {
    return res.status(404).json({ error: `O arquivo Excel ativo não foi encontrado: ${activeExcelName}. Por favor, envie uma nova planilha.` });
  }
  
  const outputDir = path.join(__dirname, 'temp_prints');
  
  const psCommand = `powershell -ExecutionPolicy Bypass -File "${path.join(__dirname, 'generate_prints.ps1')}" -ExcelPath "${excelPath}" -OutputDir "${outputDir}"`;
  
  exec(psCommand, (error, stdout, stderr) => {
    if (error) {
      console.error('PowerShell error:', error);
      console.error('stderr:', stderr);
      return res.status(500).json({ error: 'Falha ao processar Excel: ' + error.message });
    }
    
    const match = stdout.match(/JSON_START\r?\n([\s\S]*?)\r?\nJSON_END/);
    if (match) {
      try {
        const drivers = JSON.parse(match[1]);
        // Map absolute local path to web accessible relative URL
        drivers.forEach(d => {
          if (d.ImagePath) {
            const filename = path.basename(d.ImagePath);
            d.ImageUrl = `/temp_prints/${filename}`;
          }
        });
        res.json({ success: true, drivers });
      } catch (e) {
        console.error('Failed to parse JSON output from PowerShell:', e);
        res.status(500).json({ error: 'Saída JSON inválida do PowerShell' });
      }
    } else {
      res.status(500).json({ error: 'Nenhuma saída JSON encontrada na execução do PowerShell' });
    }
  });
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
