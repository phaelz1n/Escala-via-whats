// app.js
document.addEventListener('DOMContentLoaded', () => {
    // State Variables
    let drivers = [];
    let isSendingMass = false;
    let isPaused = false;
    let currentSendIndex = 0;
    let selectedDriverIds = [];
    let whatsappStatus = 'disconnected';
    let statusCheckInterval = null;

    // DOM Elements
    const btnOpenWa = document.getElementById('btn-open-wa');
    const btnCloseWa = document.getElementById('btn-close-wa');
    const statusDot = document.getElementById('status-dot');
    const statusText = document.getElementById('status-text');
    const waConnectionTip = document.getElementById('wa-connection-tip');
    const waQrContainer = document.getElementById('wa-qr-container');
    const waQrImg = document.getElementById('wa-qr-img');
    
    const inputDdd = document.getElementById('input-ddd');
    const inputTemplate = document.getElementById('input-template');
    
    const labelExcelName = document.getElementById('label-excel-name');
    const btnSelectExcel = document.getElementById('btn-select-excel');
    const inputExcelFile = document.getElementById('input-excel-file');
    const btnLoadExcel = document.getElementById('btn-load-excel');
    const loadingOverlay = document.getElementById('loading-overlay');
    const loadingTitle = document.getElementById('loading-title');
    const loadingText = document.getElementById('loading-text');

    const driversSection = document.getElementById('drivers-section');
    const driverCount = document.getElementById('driver-count');
    const emptyState = document.getElementById('drivers-empty-state');
    const tableWrapper = document.getElementById('drivers-table-wrapper');
    const tableBody = document.getElementById('drivers-table-body');
    const selectAllDrivers = document.getElementById('select-all-drivers');
    
    const massActions = document.getElementById('mass-actions');
    const selectedCountSpan = document.getElementById('selected-count');
    const btnSendSelected = document.getElementById('btn-send-selected');
    const btnPauseSend = document.getElementById('btn-pause-send');
    
    const globalProgressBar = document.getElementById('global-progress-bar');
    const progressStatusLabel = document.getElementById('progress-status-label');
    const progressPercent = document.getElementById('progress-percent');
    const progressFillBar = document.getElementById('progress-fill-bar');

    const imageLightbox = document.getElementById('image-lightbox');
    const lightboxTitle = document.getElementById('lightbox-title');
    const lightboxImg = document.getElementById('lightbox-img');
    const btnCloseLightbox = document.getElementById('btn-close-lightbox');

    // ----------------------------------------------------
    // WhatsApp Connections & Monitoring
    // ----------------------------------------------------
    
    // Poll WhatsApp Web state
    async function updateWhatsAppStatus() {
        try {
            const res = await fetch('/api/whatsapp/status');
            const data = await res.json();
            whatsappStatus = data.status;
            
            // UI Update based on status
            statusDot.className = 'status-indicator';
            
            if (whatsappStatus === 'connected') {
                statusDot.classList.add('status-connected');
                statusText.textContent = 'WhatsApp Conectado';
                btnOpenWa.disabled = true;
                btnCloseWa.disabled = false;
                waConnectionTip.classList.add('hide');
                waQrContainer.classList.add('hide');
                enableSendingUI(true);
            } else if (whatsappStatus === 'qr_ready') {
                statusDot.classList.add('status-loading');
                statusText.textContent = 'Aguardando QR Code';
                btnOpenWa.disabled = true;
                btnCloseWa.disabled = false;
                waConnectionTip.classList.remove('hide');
                waQrContainer.classList.remove('hide');
                enableSendingUI(false);
                
                // Fetch the QR code image
                try {
                    const qrRes = await fetch('/api/whatsapp/qr');
                    const qrData = await qrRes.json();
                    if (qrData.qr) {
                        waQrImg.src = qrData.qr;
                    }
                } catch (qrErr) {
                    console.error('Error fetching QR image:', qrErr);
                }
            } else if (whatsappStatus === 'loading') {
                statusDot.classList.add('status-loading');
                statusText.textContent = 'Carregando WhatsApp...';
                btnOpenWa.disabled = true;
                btnCloseWa.disabled = false;
                waConnectionTip.classList.add('hide');
                waQrContainer.classList.add('hide');
                enableSendingUI(false);
            } else { // disconnected
                statusDot.classList.add('status-disconnected');
                statusText.textContent = 'WhatsApp Desconectado';
                btnOpenWa.disabled = false;
                btnCloseWa.disabled = true;
                waConnectionTip.classList.add('hide');
                waQrContainer.classList.add('hide');
                enableSendingUI(false);
            }
        } catch (e) {
            console.error('Error polling WhatsApp status:', e);
        }
    }

    function enableSendingUI(enabled) {
        // Toggle individual send buttons
        document.querySelectorAll('.btn-send-individual').forEach(btn => {
            btn.disabled = !enabled;
        });
        // Toggle mass send buttons
        btnSendSelected.disabled = !enabled;
    }

    // Start WhatsApp polling
    updateWhatsAppStatus();
    statusCheckInterval = setInterval(updateWhatsAppStatus, 4000);

    // Open WhatsApp Web
    btnOpenWa.addEventListener('click', async () => {
        try {
            showLoadingOverlay('Iniciando WhatsApp', 'Iniciando o WhatsApp Web em segundo plano no servidor. Por favor, aguarde...');
            const res = await fetch('/api/whatsapp/start', { method: 'POST' });
            await res.json();
            updateWhatsAppStatus();
        } catch (e) {
            alert('Falha ao abrir WhatsApp: ' + e.message);
        } finally {
            hideLoadingOverlay();
        }
    });

    // Close WhatsApp Connection
    btnCloseWa.addEventListener('click', async () => {
        try {
            showLoadingOverlay('Fechando Conexão', 'Encerrando o navegador e limpando os processos do Chrome...');
            const res = await fetch('/api/whatsapp/stop', { method: 'POST' });
            await res.json();
            updateWhatsAppStatus();
        } catch (e) {
            alert('Erro ao fechar conexão: ' + e.message);
        } finally {
            hideLoadingOverlay();
        }
    });

    // Fetch active excel file on load
    async function fetchActiveExcelName() {
        try {
            const res = await fetch('/api/active-excel');
            const data = await res.json();
            if (data.filename) {
                labelExcelName.textContent = data.filename;
            }
        } catch (e) {
            console.error('Error fetching active excel name:', e);
            labelExcelName.textContent = 'Erro ao carregar';
        }
    }
    fetchActiveExcelName();

    // Select Excel File trigger
    btnSelectExcel.addEventListener('click', () => {
        inputExcelFile.click();
    });

    // Handle Excel file selection & upload
    inputExcelFile.addEventListener('change', async () => {
        const file = inputExcelFile.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = async (e) => {
            const base64Content = e.target.result.split(',')[1];
            showLoadingOverlay('Enviando Planilha', `Salvando ${file.name} no servidor e definindo como escala ativa...`);
            try {
                const res = await fetch('/api/upload', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ filename: file.name, base64: base64Content })
                });
                const data = await res.json();
                if (data.success) {
                    labelExcelName.textContent = data.filename;
                    inputExcelFile.value = '';
                    // Automatically trigger processing
                    btnLoadExcel.click();
                } else {
                    alert('Erro ao enviar planilha: ' + (data.error || 'Erro desconhecido'));
                }
            } catch (err) {
                alert('Falha na conexão de rede: ' + err.message);
            } finally {
                hideLoadingOverlay();
            }
        };
        reader.readAsDataURL(file);
    });

    // ----------------------------------------------------
    // Load Excel & Driver Management
    // ----------------------------------------------------
    
    btnLoadExcel.addEventListener('click', async () => {
        showLoadingOverlay('Processando Planilha', 'Analisando os dados da planilha de escalas e gerando as imagens dos motoristas em segundo plano. Por favor, aguarde...');
        try {
            const selectedTab = document.getElementById('select-excel-tab').value;
            const res = await fetch(`/api/drivers?tab=${selectedTab}`);
            const data = await res.json();
            
            if (data.success && data.drivers) {
                drivers = data.drivers;
                renderDriversTable();
            } else {
                alert('Erro ao carregar: ' + (data.error || 'Erro desconhecido'));
            }
        } catch (e) {
            alert('Erro de rede ao carregar planilhas: ' + e.message);
        } finally {
            hideLoadingOverlay();
        }
    });

    function renderDriversTable() {
        tableBody.innerHTML = '';
        selectedDriverIds = [];
        
        if (drivers.length === 0) {
            emptyState.classList.remove('hide');
            tableWrapper.classList.add('hide');
            massActions.style.display = 'none';
            driverCount.textContent = '0 total';
            return;
        }

        emptyState.classList.add('hide');
        tableWrapper.classList.remove('hide');
        massActions.style.display = 'flex';
        
        const activeDrivers = drivers.filter(d => d.HasScale).length;
        driverCount.textContent = `${drivers.length} total (${activeDrivers} com escalas)`;

        drivers.forEach((driver, index) => {
            const tr = document.createElement('tr');
            tr.setAttribute('data-id', index);
            
            // Status markup
            const scaleBadge = driver.HasScale 
                ? '<span class="badge badge-success"><i class="fa-solid fa-circle-check"></i> Sim</span>' 
                : '<span class="badge badge-neutral"><i class="fa-solid fa-circle-minus"></i> Não</span>';

            const viewButton = driver.HasScale 
                ? `<button class="btn-view-print" data-url="${driver.ImageUrl}?t=${Date.now()}" data-name="${driver.Name}">
                     <i class="fa-regular fa-image"></i> Ver Print
                   </button>` 
                : '<span class="text-muted">-</span>';

            const checkboxMarkup = driver.HasScale
                ? `<label class="checkbox-container">
                     <input type="checkbox" class="driver-checkbox" data-id="${index}">
                     <span class="checkmark"></span>
                   </label>`
                : `<label class="checkbox-container">
                     <input type="checkbox" disabled>
                     <span class="checkmark" style="opacity: 0.3;"></span>
                   </label>`;

            const sendStatusId = `send-status-${index}`;
            const sendStatusMarkup = `<span class="badge badge-neutral" id="${sendStatusId}">Pendente</span>`;

            const sendBtnMarkup = driver.HasScale
                ? `<button class="btn btn-primary btn-success btn-send-individual" data-id="${index}" ${whatsappStatus !== 'connected' ? 'disabled' : ''}>
                     <i class="fa-solid fa-paper-plane"></i> Enviar
                   </button>`
                : `<button class="btn btn-secondary" disabled>
                     <i class="fa-solid fa-ban"></i>
                   </button>`;

            tr.innerHTML = `
                <td>${checkboxMarkup}</td>
                <td style="font-weight: 500;">${driver.Name}</td>
                <td>${driver.Phone || '<span class="text-muted">Sem telefone</span>'}</td>
                <td>${scaleBadge}</td>
                <td>${viewButton}</td>
                <td>${sendStatusMarkup}</td>
                <td style="text-align: right;">${sendBtnMarkup}</td>
            `;

            tableBody.appendChild(tr);
        });

        // Event listeners for individual send
        document.querySelectorAll('.btn-send-individual').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const id = btn.getAttribute('data-id');
                await sendIndividualMessage(id);
            });
        });

        // Event listeners for visual preview
        document.querySelectorAll('.btn-view-print').forEach(btn => {
            btn.addEventListener('click', () => {
                const url = btn.getAttribute('data-url');
                const name = btn.getAttribute('data-name');
                openLightbox(url, name);
            });
        });

        // Checkbox events
        document.querySelectorAll('.driver-checkbox').forEach(cb => {
            cb.addEventListener('change', () => {
                const id = parseInt(cb.getAttribute('data-id'));
                if (cb.checked) {
                    selectedDriverIds.push(id);
                } else {
                    selectedDriverIds = selectedDriverIds.filter(item => item !== id);
                }
                updateSelectedCount();
            });
        });

        selectAllDrivers.checked = false;
        updateSelectedCount();
    }

    // Handle Select All Checkbox
    selectAllDrivers.addEventListener('change', () => {
        const checkBoxes = document.querySelectorAll('.driver-checkbox');
        selectedDriverIds = [];
        
        checkBoxes.forEach(cb => {
            cb.checked = selectAllDrivers.checked;
            if (cb.checked) {
                const id = parseInt(cb.getAttribute('data-id'));
                selectedDriverIds.push(id);
            }
        });
        updateSelectedCount();
    });

    function updateSelectedCount() {
        selectedCountSpan.textContent = selectedDriverIds.length;
        btnSendSelected.disabled = selectedDriverIds.length === 0 || whatsappStatus !== 'connected';
    }

    // ----------------------------------------------------
    // Message Dispatch Queue
    // ----------------------------------------------------
    
    // Send message to one driver
    async function sendIndividualMessage(index) {
        const driver = drivers[index];
        const statusBadge = document.getElementById(`send-status-${index}`);
        const individualBtn = document.querySelector(`.btn-send-individual[data-id="${index}"]`);
        
        if (!driver || !driver.HasScale) return false;

        // Set status UI to Sending
        statusBadge.className = 'badge badge-warning';
        statusBadge.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Enviando...';
        if (individualBtn) individualBtn.disabled = true;

        try {
            // Replace template placeholder
            let msgText = inputTemplate.value.replace(/{motorista}/gi, driver.Name);
            
            const payload = {
                name: driver.Name,
                phone: driver.Phone,
                text: msgText,
                imagePath: driver.ImagePath
            };

            const res = await fetch('/api/whatsapp/send', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });
            
            const result = await res.json();
            
            if (result.success) {
                statusBadge.className = 'badge badge-success';
                statusBadge.innerHTML = '<i class="fa-solid fa-check"></i> Enviado';
                return true;
            } else {
                throw new Error(result.error || 'Erro desconhecido de envio.');
            }
        } catch (e) {
            statusBadge.className = 'badge badge-danger';
            statusBadge.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Falhou';
            statusBadge.title = e.message;
            if (individualBtn) individualBtn.disabled = false;
            return false;
        }
    }

    // Mass sending routine
    btnSendSelected.addEventListener('click', async () => {
        if (selectedDriverIds.length === 0) return;
        
        isSendingMass = true;
        isPaused = false;
        currentSendIndex = 0;
        
        // Show progress bar
        globalProgressBar.classList.remove('hide');
        btnPauseSend.classList.remove('hide');
        btnSendSelected.disabled = true;
        btnLoadExcel.disabled = true;
        selectAllDrivers.disabled = true;
        document.querySelectorAll('.driver-checkbox').forEach(cb => cb.disabled = true);
        
        btnPauseSend.innerHTML = '<i class="fa-solid fa-pause"></i> Pausar Envio';

        runMassSendQueue();
    });

    async function runMassSendQueue() {
        if (!isSendingMass) return;

        if (currentSendIndex >= selectedDriverIds.length) {
            // Finished
            progressStatusLabel.textContent = 'Envios concluídos!';
            cleanupMassSendUI();
            alert('Envio em massa de escalas finalizado!');
            return;
        }

        if (isPaused) {
            progressStatusLabel.textContent = 'Envio pausado...';
            return;
        }

        const driverIndex = selectedDriverIds[currentSendIndex];
        const driver = drivers[driverIndex];

        // Update progress bar labels
        progressStatusLabel.textContent = `Enviando para ${driver.Name} (${currentSendIndex + 1}/${selectedDriverIds.length})...`;
        
        // Update percentages
        const percent = Math.round((currentSendIndex / selectedDriverIds.length) * 100);
        progressPercent.textContent = `${percent}%`;
        progressFillBar.style.width = `${percent}%`;

        // Send individual message
        await sendIndividualMessage(driverIndex);
        
        // Move to next driver
        currentSendIndex++;
        
        // Final progress update
        const finalPercent = Math.round((currentSendIndex / selectedDriverIds.length) * 100);
        progressPercent.textContent = `${finalPercent}%`;
        progressFillBar.style.width = `${finalPercent}%`;

        // Introduce a subtle natural delay between messages
        setTimeout(runMassSendQueue, 2000);
    }

    btnPauseSend.addEventListener('click', () => {
        if (isPaused) {
            isPaused = false;
            btnPauseSend.innerHTML = '<i class="fa-solid fa-pause"></i> Pausar Envio';
            progressStatusLabel.textContent = 'Retomando envios...';
            runMassSendQueue();
        } else {
            isPaused = true;
            btnPauseSend.innerHTML = '<i class="fa-solid fa-play"></i> Continuar';
            progressStatusLabel.textContent = 'Pausando envio...';
        }
    });

    function cleanupMassSendUI() {
        isSendingMass = false;
        isPaused = false;
        btnPauseSend.classList.add('hide');
        btnLoadExcel.disabled = false;
        selectAllDrivers.disabled = false;
        
        // Reactivate checkboxes
        document.querySelectorAll('.driver-checkbox').forEach(cb => {
            const id = parseInt(cb.getAttribute('data-id'));
            const driver = drivers[id];
            if (driver && driver.HasScale) {
                cb.disabled = false;
            }
        });
        
        updateSelectedCount();
    }

    // ----------------------------------------------------
    // Lightbox & Overlays
    // ----------------------------------------------------
    
    function openLightbox(url, name) {
        const baseUrl = url.split('?')[0];
        lightboxImg.src = `${baseUrl}?t=${Date.now()}`;
        lightboxTitle.textContent = `Escala de Serviço - ${name}`;
        imageLightbox.classList.remove('hide');
    }

    function closeLightbox() {
        imageLightbox.classList.add('hide');
        lightboxImg.src = '';
    }

    btnCloseLightbox.addEventListener('click', closeLightbox);
    
    imageLightbox.addEventListener('click', (e) => {
        if (e.target === imageLightbox) {
            closeLightbox();
        }
    });

    function showLoadingOverlay(title, text) {
        loadingTitle.textContent = title;
        loadingText.textContent = text;
        loadingOverlay.classList.remove('hide');
    }

    function hideLoadingOverlay() {
        loadingOverlay.classList.add('hide');
    }
});
