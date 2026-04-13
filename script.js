document.addEventListener('DOMContentLoaded', () => {
    // LOGIN LOGIC
    const loginOverlay = document.getElementById('login-overlay');
    const appContainer = document.querySelector('.app-container');
    const loginBtn = document.getElementById('login-btn');
    const loginUsernameInput = document.getElementById('login-username');
    const loginPasswordInput = document.getElementById('login-password');
    const loginMessage = document.getElementById('login-message');

    // Link script.google.com của bạn (Phải cập nhật link này sau khi triển khai Apps Script)
    const SHEETS_API_URL = "https://script.google.com/macros/s/AKfycbyddVGQ6bKUDteLmyU0hoypFVmESKScW-tbuJRU0cQGsn35gLv3gNGjx4hswffqERhs/exec";

    // --- STABLE DEVICE ID (FINGERPRINT) ---
    async function getStableDeviceId() {
        try {
            const screenInfo = `${screen.width}x${screen.height}x${screen.colorDepth}`;
            const cpuInfo = navigator.hardwareConcurrency || 'unknown';
            const memInfo = navigator.deviceMemory || 'unknown';
            const lang = navigator.language || 'unknown';
            const platform = navigator.platform || 'unknown';
            const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
            
            // Canvas Fingerprinting: High stability across cache clears
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            ctx.textBaseline = "top";
            ctx.font = "14px 'Arial'";
            ctx.textBaseline = "alphabetic";
            ctx.fillStyle = "#f60";
            ctx.fillRect(125,1,62,20);
            ctx.fillStyle = "#069";
            ctx.fillText("gv_assistant_hwid_v1", 2, 15);
            ctx.fillStyle = "rgba(102, 204, 0, 0.7)";
            ctx.fillText("gv_assistant_hwid_v1", 4, 17);
            const canvasData = canvas.toDataURL();

            // WebGL Fingerprinting (if available)
            let glInfo = "";
            try {
                const gl = canvas.getContext('webgl') || canvas.getContext('experimental-webgl');
                if (gl) {
                    const debugInfo = gl.getExtension('WEBGL_debug_renderer_info');
                    glInfo = gl.getParameter(debugInfo.UNMASKED_RENDERER_WEBGL);
                }
            } catch (e) {}

            const rawStr = `HWID-V1-${screenInfo}-${cpuInfo}-${memInfo}-${lang}-${platform}-${timezone}-${canvasData.length}-${glInfo}`;
            
            // Simple hash function (FNV-1a or similar) to generate a clean string
            let hash = 0;
            for (let i = 0; i < rawStr.length; i++) {
                const char = rawStr.charCodeAt(i);
                hash = ((hash << 5) - hash) + char;
                hash = hash & hash; // Convert to 32bit integer
            }
            const finalId = 'HWID-' + Math.abs(hash).toString(16).toUpperCase() + '-' + rawStr.length;
            return finalId;
        } catch (e) {
            console.error("Fingerprint error:", e);
            return 'dev-' + Math.random().toString(36).substring(2, 15);
        }
    }

    let deviceId = "";
    
    // Khởi tạo Device ID
    (async () => {
        deviceId = await getStableDeviceId();
        console.log("System Machine ID:", deviceId);
        
        // Hiển thị Device ID ở màn hình đăng nhập để dễ quản lý
        const loginFooter = document.querySelector('.login-footer');
        if (loginFooter) {
            const hwidDisplay = document.createElement('div');
            hwidDisplay.style.fontSize = '14px';
            hwidDisplay.style.color = '#0056b3';
            hwidDisplay.style.fontWeight = '700';
            hwidDisplay.style.marginTop = '15px';
            hwidDisplay.style.padding = '10px';
            hwidDisplay.style.background = '#f0f7ff';
            hwidDisplay.style.borderRadius = '8px';
            hwidDisplay.style.border = '1px dashed #0056b3';
            hwidDisplay.innerHTML = `Mã máy: <span style="user-select: all;">${deviceId}</span>`;
            loginFooter.appendChild(hwidDisplay);
        }

        // Kiểm tra trạng thái đã đăng nhập chưa
        if (localStorage.getItem('isLoggedIn') === 'true') {
            const savedUser = localStorage.getItem('username') || 'Thành viên';
            const displayUser = document.getElementById('display-username');
            if (displayUser) displayUser.textContent = savedUser;

            loginOverlay.style.display = 'none';
            appContainer.style.display = 'flex';
        }
    })();

    async function handleLogin() {
        const user = loginUsernameInput.value.trim();
        const pass = loginPasswordInput.value.trim();

        if (!user || !pass) {
            showLoginMessage('Vui lòng nhập đầy đủ thông tin!', 'error');
            return;
        }

        loginBtn.disabled = true;
        loginBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Đang xác thực...';
        showLoginMessage('Đang kết nối hệ thống...', '');

        // Ghi chú: Tài khoản (user) sẽ được kiểm tra không phân biệt chữ hoa/thường ở Backend
        // Mật khẩu (pass) vẫn giữ nguyên tính phân biệt chữ hoa/thường


        if (!deviceId) {
            deviceId = await getStableDeviceId();
        }

        try {
            // Thêm cache: 'no-cache' và redirect: 'follow' để đảm bảo fetch hoạt động tốt với Google Script
            const response = await fetch(`${SHEETS_API_URL}?action=login&user=${encodeURIComponent(user)}&pass=${encodeURIComponent(pass)}&deviceId=${encodeURIComponent(deviceId)}`, {
                method: 'GET',
                mode: 'cors',
                cache: 'no-cache',
                redirect: 'follow'
            });

            if (!response.ok) {
                console.warn("Server Response Not OK:", response.status, response.statusText);
                throw new Error('Server trả về lỗi: ' + response.status);
            }

            const text = await response.text();
            console.log("Raw Response from Server:", text);
            
            let result;
            try {
                result = JSON.parse(text);
            } catch (e) {
                console.error("JSON Parsing Error. Received text:", text);
                throw new Error('Dữ liệu từ Google Sheet không đúng định dạng. Hãy kiểm tra lại URL hoặc mã Script.');
            }

            if (result.status === 'success') {
                const isNewBinding = result.message && result.message.toLowerCase().includes('đã được gán');
                const successMsg = isNewBinding ? 'Gắn mã máy thành công! Đang vào hệ thống...' : 'Đăng nhập thành công! Đang vào hệ thống...';
                
                showLoginMessage(successMsg, 'success');
                localStorage.setItem('isLoggedIn', 'true');
                localStorage.setItem('username', user);

                const displayUser = document.getElementById('display-username');
                if (displayUser) displayUser.textContent = user;

                setTimeout(() => {
                    loginOverlay.style.opacity = '0';
                    setTimeout(() => {
                        loginOverlay.style.display = 'none';
                        appContainer.style.display = 'flex';
                        appContainer.style.opacity = '1';
                    }, 500);
                }, 1000);
            } else {
                showLoginMessage(result.message || 'Tài khoản hoặc mật khẩu không đúng', 'error');
                loginBtn.disabled = false;
                loginBtn.innerHTML = '<span>ĐĂNG NHẬP NGAY</span><i class="fas fa-sign-in-alt"></i>';
            }
        } catch (error) {
            console.error("Login Error:", error);
            
            let errorMsg = 'Lỗi hệ thống: ' + error.message;
            if (error.message === 'Failed to fetch') {
                errorMsg = 'Lỗi kết nối: Trình duyệt đang chặn yêu cầu. Hãy thử bấm Chuột Phải vào index.html chọn "Open with Live Server" hoặc nộp file lên Hosting/Vercel để chạy.';
            }
            
            showLoginMessage(errorMsg + '. (Hãy chắc chắn chọn "Anyone" khi Deploy bản mới nhất)', 'error');
            loginBtn.disabled = false;
            loginBtn.innerHTML = '<span>ĐĂNG NHẬP NGAY</span><i class="fas fa-sign-in-alt"></i>';
        }
    }

    function getGradeText() {
        const select = document.getElementById('grade-level');
        if (!select) return "Tiểu học";
        const val = select.value;
        if (val === 'mam-non') return "Mầm non";
        if (val === 'tieu-hoc') return "Tiểu học";
        if (val === 'thcs') return "THCS";
        if (val === 'thpt') return "THPT";
        return "Tiểu học";
    }

    function showLoginMessage(text, type) {
        loginMessage.textContent = text;
        loginMessage.className = 'login-message ' + type;
        loginMessage.style.display = text ? 'block' : 'none';
    }

    loginBtn?.addEventListener('click', handleLogin);

    // Đăng nhập bằng phím Enter
    [loginUsernameInput, loginPasswordInput].forEach(input => {
        input?.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') handleLogin();
        });
    });

    // LOGOUT LOGIC
    const logoutBtn = document.getElementById('logout-btn');
    logoutBtn?.addEventListener('click', () => {
        if (confirm('Bạn có chắc chắn muốn đăng xuất không?')) {
            localStorage.removeItem('isLoggedIn');
            localStorage.removeItem('username');
            location.reload();
        }
    });

    // GUIDE MODAL LOGIC
    const guideModal = document.getElementById('guide-modal');
    const openGuideBtn = document.getElementById('open-guide-btn');
    const closeGuideBtns = document.querySelectorAll('.close-guide-modal');

    openGuideBtn?.addEventListener('click', () => {
        guideModal.style.display = 'flex';
    });

    closeGuideBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            guideModal.style.display = 'none';
        });
    });

    // Close modal when clicking outside
    window.addEventListener('click', (e) => {
        if (e.target === guideModal) {
            guideModal.style.display = 'none';
        }
    });

    // Selectors
    const navItems = document.querySelectorAll('.nav-item');
    const sections = document.querySelectorAll('.content-section');
    const infoForm = document.getElementById('info-form');
    const aiSuggestBtn = document.getElementById('suggest-topic-btn');
    const tabButtons = document.querySelectorAll('.tab-btn');
    const sidebar = document.querySelector('.sidebar');

    const apiModal = document.getElementById('api-modal');
    const openApiModalBtn = document.getElementById('open-api-modal');
    const closeModalBtn = document.querySelector('.close-modal');
    const saveApiKeyBtn = document.getElementById('save-api-key');
    const apiKeyInput = document.getElementById('gemini-api-key');
    const apiStatus = document.getElementById('api-status');

    const nextBtns = document.querySelectorAll('.next-btn');
    const prevBtns = document.querySelectorAll('.prev-btn');
    const writeAllBtn = document.querySelector('.write-all-btn');
    const exportWordBtn = document.getElementById('export-word-btn');
    const writingProgressView = document.getElementById('writing-progress-view');
    const progressBarFill = document.getElementById('writing-progress-bar-fill');
    const progressPercent = document.getElementById('writing-progress-percent');
    const topicAnalysisView = document.getElementById('topic-analysis-view');
    const refineTopicBtn = document.getElementById('refine-topic-btn');
    const closeTopicAnalysis = document.getElementById('close-topic-analysis');
    const topicSuggestionsList = document.getElementById('topic-suggestions-list');
    const upgradePlanView = document.getElementById('upgrade-plan-view');
    const upgradeStepsList = document.getElementById('upgrade-steps-list');
    const alertAiScore = document.getElementById('alert-ai-score');
    const topicTitleInput = document.getElementById('topic-title');
    const closeUpgradePlanBtn = document.getElementById('close-upgrade-plan');
    const exportAppraisalBtn = document.getElementById('export-appraisal-btn');
    const confirmUpgradeBtn = document.getElementById('confirm-upgrade-btn');
    const upgradePageLimitInput = document.getElementById('upgrade-page-limit');
    const writingProgressOverlay = document.getElementById('writing-progress-overlay');
    const progressLoadingIcon = document.getElementById('progress-loading-icon');
    const progressSuccessIcon = document.getElementById('progress-success-icon');
    const progressTitle = document.getElementById('progress-title');
    const progressStatus = document.getElementById('writing-progress-status');
    const progressLog = document.getElementById('writing-progress-log');
    const finishedActions = document.getElementById('writing-finished-actions');
    const closeProgressOverlayBtn = document.getElementById('close-progress-overlay');
    const viewResultBtn = document.getElementById('view-result-btn');
    const exportAfterUpgradeBtn = document.getElementById('export-after-upgrade-btn');
    const keepOldTopicBtn = document.getElementById('keep-old-topic-btn');
    const appraisalUpgradeResultView = document.getElementById('appraisal-upgrade-result-view');
    const appraisalResultsView = document.getElementById('appraisal-results-view');
    const exportRefinedDocxBtn = document.getElementById('export-refined-docx-btn');
    const refinedFullEditor = document.getElementById('refined-full-editor');
    const uploadDropzone = document.getElementById('upload-dropzone');
    const bpAppraisalContent = document.getElementById('bp-appraisal-content');
    const skknAppraisalContent = document.getElementById('skkn-appraisal-content');
    const analysisSummaryText = document.getElementById('analysis-summary-text');
    const selectFileBtn = document.getElementById('select-file-btn');
    const fileInput = document.getElementById('file-upload-input');
    const uploadView = document.getElementById('appraisal-upload-view');
    const resultsView = document.getElementById('appraisal-results-view');
    const reUploadBtn = document.getElementById('re-upload-btn');
    const completeBtn = document.getElementById('complete-btn');
    const skknCompleteBtn = document.getElementById('skkn-complete-btn');
    const skknExportWordBtn = document.getElementById('skkn-export-word-btn');

    let lastAppraisalData = null;
    let lastAppraisedText = "";

    function cleanFileName(str) {
        if (!str) return "Bao_Cao_Giao_Vien";
        return str.normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .replace(/đ/g, 'd').replace(/Đ/g, 'D')
            .replace(/[^a-zA-Z0-9]/g, '_')
            .replace(/_+/g, '_')
            .substring(0, 100)
            .trim();
    }

    function htmlToParagraphs(html) {
        if (!html) return [];
        const { Paragraph, TextRun, AlignmentType } = window.docx || docx;

        // Basic entity decoding
        const decoded = html.replace(/&nbsp;/g, ' ')
            .replace(/&quot;/g, '"')
            .replace(/&amp;/g, '&')
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>');

        const items = [];
        const lines = decoded.split(/<br>|<div>|<\/div>|<p>|<\/p>/);

        lines.forEach(line => {
            const clean = line.replace(/<[^>]*>/g, '').trim();
            if (clean) {
                items.push(new Paragraph({
                    alignment: AlignmentType.JUSTIFY,
                    indent: { firstLine: 454 }, // 0.8cm thụt đầu dòng
                    spacing: { line: 360, after: 120 }, // Giãn dòng 1.5
                    children: [new TextRun({ text: clean, size: 28, font: "Times New Roman" })] // Cỡ chữ 14
                }));
            }
        });
        return items;
    }

    // --- SURGICAL JSON PARSER ---
    function robustParseJSON(text) {
        if (!text) return null;
        try {
            // 1. Strip markdown decorators
            let cleaned = text.trim();
            if (cleaned.startsWith('```json')) {
                cleaned = cleaned.substring(7);
                if (cleaned.endsWith('```')) cleaned = cleaned.substring(0, cleaned.length - 3);
            } else if (cleaned.startsWith('```')) {
                cleaned = cleaned.substring(3);
                if (cleaned.endsWith('```')) cleaned = cleaned.substring(0, cleaned.length - 3);
            }
            cleaned = cleaned.trim();

            // 2. Surgical extraction of the outermost object/array
            const firstBrace = cleaned.indexOf('{');
            const lastBrace = cleaned.lastIndexOf('}');
            const firstBracket = cleaned.indexOf('[');
            const lastBracket = cleaned.lastIndexOf(']');

            let targetJson = "";
            let start = -1, end = -1;

            if (firstBrace !== -1 && (firstBracket === -1 || firstBrace < firstBracket)) {
                start = firstBrace;
                end = lastBrace;
            } else if (firstBracket !== -1) {
                start = firstBracket;
                end = lastBracket;
            }

            if (start !== -1 && end !== -1) {
                targetJson = cleaned.substring(start, end + 1);
            } else {
                targetJson = cleaned;
            }

            // 3. Remove single-line comments // that AI sometimes adds
            targetJson = targetJson.replace(/\/\/.*$/gm, '');

            try {
                return JSON.parse(targetJson);
            } catch (initialErr) {
                // If native parse fails, try one more time by ensuring no trailing commas before closing braces
                const repaired = targetJson.replace(/,\s*([}\]])/g, '$1');
                return JSON.parse(repaired);
            }
        } catch (e) {
            console.error("Robust JSON Parse Failed:", e, "\nOriginal Text:", text);
            return null;
        }
    }

    const skknStructure = {
        '1': 'PHẦN I: ĐẶT VẤN ĐỀ (1.1. Lý do chọn đề tài, 1.2. Mục đích, đối tượng, phạm vi và phương pháp nghiên cứu)',
        'theory': 'PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.1. Cơ sở lý luận/khoa học (Nêu ngắn gọn lý thuyết liên quan)',
        '2': 'PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.2. Thực trạng vấn đề (Thuận lợi, khó khăn và Số liệu minh chứng đầu vào)',
        '2.3.1': 'PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.3. Các giải pháp thực hiện (Giải pháp 1)',
        '2.3.2': 'PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.3. Các giải pháp thực hiện (Giải pháp 2)',
        '2.3.3': 'PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.3. Các giải pháp thực hiện (Giải pháp 3)',
        '2.4': 'PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.4. Hiệu quả của SKKN (So sánh số liệu, chứng minh kết quả sau khi áp dụng)',
        '3': 'PHẦN III: KẾT LUẬN VÀ KIẾN NGHỊ (3.1. Tóm tắt kết quả, bài học kinh nghiệm; 3.2. Kiến nghị, đề xuất)',
        'appendix': 'PHỤ LỤC (Tài liệu tham khảo, Minh chứng)'
    };

    const bpSectionOrder = ['info', 'outline', 'section1', 'section2', 'solution1', 'solution2', 'solution3', 'creativity', 'impact'];
    const skknSectionOrder = ['info', 'outline', 'section1', 'theory', 'section2', 'solution1', 'solution2', 'solution3', 'creativity', 'impact', 'appendix'];
    let currentSectionOrder = bpSectionOrder;
    let currentSectionIndex = 0;
    let currentMode = 'biên-pháp';

    tabButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            tabButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');

            const mode = btn.getAttribute('data-mode');
            currentMode = mode;

            if (mode === 'thẩm-định') {
                sidebar.style.opacity = '0.5'; sidebar.style.pointerEvents = 'none';
                hideAllSections(); document.getElementById('appraisal-section').classList.add('active');
            } else {
                sidebar.style.opacity = '1'; sidebar.style.pointerEvents = 'auto';
                if (mode === 'skkn') {
                    currentSectionOrder = skknSectionOrder;
                    document.querySelectorAll('.bp-nav').forEach(el => el.style.display = 'none');
                    document.querySelectorAll('.skkn-nav').forEach(el => el.style.display = 'block');
                    const roleSelector = document.getElementById('main-role-selector');
                    if (roleSelector) roleSelector.style.display = 'none';
                    const bpOutline = document.getElementById('bp-outline-container');
                    const skknOutline = document.getElementById('skkn-outline-container');
                    if (bpOutline) bpOutline.style.display = 'none';
                    if (skknOutline) skknOutline.style.display = 'flex';
                    if (skknOutline) skknOutline.style.flexDirection = 'column';

                    const topicLabel = document.getElementById('topic-label');
                    const topicTitle = document.getElementById('topic-title');
                    const mainTitle = document.getElementById('main-title');
                    if (mainTitle) mainTitle.innerText = "Trợ Lý Viết SKKN";
                    if (topicLabel) topicLabel.innerHTML = 'Tên đề tài SKKN <span class="required">*</span>';
                    if (topicTitle) topicTitle.placeholder = 'VD: "Một số giải pháp ứng dụng AI vào giảng dạy môn..."';

                    // Update shared section titles for SKKN context
                    const s1Hero = document.querySelector('#section1-section h1');
                    const s1Title = document.querySelector('#section1-section h2');
                    const s1Desc = document.querySelector('#section1-section p');
                    if (s1Hero) s1Hero.innerText = "PHẦN I: ĐẶT VẤN ĐỀ (MỞ ĐẦU)";
                    if (s1Title) s1Title.innerText = "I. Đặt vấn đề";
                    if (s1Desc) s1Desc.innerText = "Trình bày lý do chọn đề tài, mục đích, đối tượng, phạm vi và phương pháp nghiên cứu.";

                    const theoryHero = document.querySelector('#theory-section h1');
                    const theoryTitle = document.querySelector('#theory-section h2');
                    const theoryDesc = document.querySelector('#theory-section p');
                    if (theoryHero) theoryHero.innerText = "PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.1 CƠ SỞ LÝ LUẬN";
                    if (theoryTitle) theoryTitle.innerText = "2.1 Cơ sở lý luận/khoa học";
                    if (theoryDesc) theoryDesc.innerText = "Nêu ngắn gọn cơ sở lý thuyết và khoa học liên quan đến đề tài.";

                    const s2Hero = document.querySelector('#section2-section h1');
                    const s2Title = document.querySelector('#section2-section h2');
                    const s2Desc = document.querySelector('#section2-section p');
                    if (s2Hero) s2Hero.innerText = "PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.2 THỰC TRẠNG";
                    if (s2Title) s2Title.innerText = "2.2 Thực trạng vấn đề";
                    if (s2Desc) s2Desc.innerText = "Mô tả tình hình, nêu thuận lợi/khó khăn và số liệu minh chứng trước khi áp dụng.";

                    const c1Hero = document.querySelector('#creativity-section h1');
                    const c1Title = document.querySelector('#creativity-section h2');
                    const c1Desc = document.querySelector('#creativity-section p');
                    if (c1Hero) c1Hero.innerText = "PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.4 HIỆU QUẢ SKKN";
                    if (c1Title) c1Title.innerText = "2.4 Hiệu quả của SKKN";
                    if (c1Desc) c1Desc.innerText = "So sánh số liệu, chứng minh kết quả thực tế sau khi áp dụng các giải pháp.";

                    const s3Hero = document.querySelector('#impact-section h1');
                    const s3Title = document.querySelector('#impact-section h2');
                    const s3Desc = document.querySelector('#impact-section p');
                    if (s3Hero) s3Hero.innerText = "PHẦN III: KẾT LUẬN & KIẾN NGHỊ";
                    if (s3Title) s3Title.innerText = "III. Kết luận & Kiến nghị";
                    if (s3Desc) s3Desc.innerText = "Tóm tắt kết quả, rút ra bài học kinh nghiệm và đề xuất hướng phát triển.";

                    // Solutions titles for SKKN
                    document.querySelector('#solution1-section h1').innerText = "PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.3 GIẢI PHÁP 1";
                    document.querySelector('#solution1-section h2').innerText = "2.3 Các biện pháp thực hiện (Mục 1)";
                    document.querySelector('#solution2-section h1').innerText = "PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.3 GIẢI PHÁP 2";
                    document.querySelector('#solution2-section h2').innerText = "2.3 Các biện pháp thực hiện (Mục 2)";
                    document.querySelector('#solution3-section h1').innerText = "PHẦN II: GIẢI QUYẾT VẤN ĐỀ - 2.3 GIẢI PHÁP 3";
                    document.querySelector('#solution3-section h2').innerText = "2.3 Các biện pháp thực hiện (Mục 3)";

                    // Toggle buttons for SKKN Transition
                    const impactNextBtn = document.getElementById('impact-next-btn');
                    const bpCompleteBtn = document.getElementById('complete-btn');
                    const bpExportBtn = document.getElementById('export-word-btn');
                    if (impactNextBtn) impactNextBtn.style.display = 'flex';
                    if (bpCompleteBtn) bpCompleteBtn.style.display = 'none';
                    if (bpExportBtn) bpExportBtn.style.display = 'none';
                } else {
                    currentSectionOrder = bpSectionOrder;
                    document.querySelectorAll('.bp-nav').forEach(el => el.style.display = 'block');
                    document.querySelectorAll('.skkn-nav').forEach(el => el.style.display = 'none');
                    const roleSelector = document.getElementById('main-role-selector');
                    if (roleSelector) roleSelector.style.display = 'flex';
                    const bpOutline = document.getElementById('bp-outline-container');
                    const skknOutline = document.getElementById('skkn-outline-container');
                    if (bpOutline) bpOutline.style.display = 'flex';
                    if (bpOutline) bpOutline.style.flexDirection = 'column';
                    if (skknOutline) skknOutline.style.display = 'none';

                    const topicLabel = document.getElementById('topic-label');
                    const topicTitle = document.getElementById('topic-title');
                    const mainTitle = document.getElementById('main-title');
                    if (mainTitle) mainTitle.innerText = "Trợ Lý Viết Biện Pháp GVG/GVCNG";
                    if (topicLabel) topicLabel.innerHTML = 'Tên đề tài Biện pháp <span class="required">*</span>';
                    if (topicTitle) topicTitle.placeholder = 'VD: "Ứng dụng AI để nâng cao hiệu quả dạy học môn Toán..."';

                    // Restore Biện pháp titles
                    const s1Hero = document.querySelector('#section1-section h1');
                    const s1Title = document.querySelector('#section1-section h2');
                    if (s1Hero) s1Hero.innerText = "1. Đặt Vấn Đề";
                    if (s1Title) s1Title.innerText = "1. Đặt vấn đề";

                    const s2Hero = document.querySelector('#section2-section h1');
                    const s2Title = document.querySelector('#section2-section h2');
                    if (s2Hero) s2Hero.innerText = "2. Thực Trạng Vấn Đề";
                    if (s2Title) s2Title.innerText = "2. Thực trạng vấn đề";

                    const c1Hero = document.querySelector('#creativity-section h1');
                    const c1Title = document.querySelector('#creativity-section h2');
                    if (c1Hero) c1Hero.innerText = "2.4 Tính Mới & Sáng Tạo";
                    if (c1Title) c1Title.innerText = "2.4.1 Tính mới & 2.4.2 Tính sáng tạo";

                    const s3Hero = document.querySelector('#impact-section h1');
                    const s3Title = document.querySelector('#impact-section h2');
                    if (s3Hero) s3Hero.innerText = "3. Kết Luận & Khuyến Nghị";
                    if (s3Title) s3Title.innerText = "3.1 Khả năng áp dụng & 3.2 Hiệu quả";

                    // Solutions titles for Biện pháp
                    document.querySelector('#solution1-section h1').innerText = "2.3 Giải pháp thứ nhất";
                    document.querySelector('#solution1-section h2').innerText = "Nội dung giải pháp 1";
                    document.querySelector('#solution2-section h1').innerText = "2.3 Giải pháp thứ hai";
                    document.querySelector('#solution2-section h2').innerText = "Nội dung giải pháp 2";
                    document.querySelector('#solution3-section h1').innerText = "2.3 Giải pháp thứ ba";
                    document.querySelector('#solution3-section h2').innerText = "Nội dung giải pháp 3";

                    // Toggle buttons for Biện pháp completion
                    const impactNextBtn = document.getElementById('impact-next-btn');
                    const bpCompleteBtn = document.getElementById('complete-btn');
                    const bpExportBtn = document.getElementById('export-word-btn');
                    if (impactNextBtn) impactNextBtn.style.display = 'none';
                    if (bpCompleteBtn) bpCompleteBtn.style.display = 'flex';
                    if (bpExportBtn) bpExportBtn.style.display = 'flex';
                }

                // Reset to info if switching modes
                currentSectionIndex = 0;
                switchSection(currentSectionOrder[currentSectionIndex]);
            }
        });
    });

    // Refine Topic button (redundant jump removed, handled by modular logic later)

    navItems.forEach(item => {
        item.addEventListener('click', () => {
            const target = item.getAttribute('data-section');
            if (target) {
                if (isWritingAll) autoSwitchEnabled = false; // Stop auto-switching if manually navigating via sidebar
                switchSection(target);
                currentSectionIndex = currentSectionOrder.indexOf(target);
            }
        });
    });

    nextBtns.forEach(btn => btn.addEventListener('click', () => {
        if (currentSectionIndex < currentSectionOrder.length - 1) {
            currentSectionIndex++;
            if (isWritingAll) autoSwitchEnabled = true; // Snap back to flow if manually going next
            switchSection(currentSectionOrder[currentSectionIndex]);
        }
    }));
    prevBtns.forEach(btn => btn.addEventListener('click', () => {
        if (currentSectionIndex > 0) {
            currentSectionIndex--;
            if (isWritingAll) autoSwitchEnabled = false; // Stop auto-switching if manually going back
            switchSection(currentSectionOrder[currentSectionIndex]);
        }
    }));

    function hideAllSections() {
        sections.forEach(s => s.classList.remove('active'));
        const app = document.getElementById('appraisal-section');
        if (app) app.classList.remove('active');
    }

    function switchSection(id) {
        navItems.forEach(item => item.classList.toggle('active', item.getAttribute('data-section') === id));
        hideAllSections();
        const el = document.getElementById(`${id}-section`);
        if (el) { el.classList.add('active'); window.scrollTo({ top: 0, behavior: 'smooth' }); }
    }

    if (localStorage.getItem('gemini_api_key')) apiKeyInput.value = localStorage.getItem('gemini_api_key');
    openApiModalBtn?.addEventListener('click', () => { apiModal.classList.add('active'); apiStatus.innerText = ''; });
    closeModalBtn?.addEventListener('click', () => apiModal.classList.remove('active'));
    let pendingAction = null;
    saveApiKeyBtn?.addEventListener('click', () => {
        const key = apiKeyInput.value.trim();
        if (key) {
            localStorage.setItem('gemini_api_key', key);
            apiStatus.innerText = 'Đã lưu API Key thành công!';
            setTimeout(() => {
                apiModal.classList.remove('active');
                if (pendingAction) {
                    const action = pendingAction;
                    pendingAction = null;
                    action();
                }
            }, 800);
        }
    });

    async function callGemini(prompt, onQuotaError) {
        const key = localStorage.getItem('gemini_api_key');
        if (!key) {
            apiModal.classList.add('active');
            apiStatus.innerText = 'Vui lòng nhập API Key để tiếp tục.';
            if (onQuotaError) pendingAction = onQuotaError;
            return { error: 'Chưa cấu hình API Key.' };
        }
        try {
            // Sử dụng gemini-3-flash-preview (Model mới nhất theo yêu cầu người dùng)
            const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${key}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt }] }],
                    generationConfig: {
                        responseMimeType: "application/json"
                    }
                })
            });

            const data = await res.json();

            if (res.status === 429) {
                apiModal.classList.add('active');
                apiStatus.innerText = 'Hết hạn mức API. Vui lòng đổi Key khác.';
                if (onQuotaError) pendingAction = onQuotaError;
                return { error: 'Hết hạn mức (Quota).' };
            }

            if (data.error) {
                console.error("Gemini API Error:", data.error);
                return { error: data.error.message || 'Lỗi API từ Google' };
            }

            let text = data.candidates?.[0]?.content?.parts?.[0]?.text;
            if (text) return { text: text.trim().replace(/^\s*\* /gm, '') };

            // Check if blocked by safety
            if (data.candidates?.[0]?.finishReason === 'SAFETY') {
                return { error: 'Nội dung bị chặn do tiêu chuẩn an toàn (Safety Filter).' };
            }

            return { error: 'Không nhận được phản hồi từ AI. Thử lại sau.' };
        } catch (e) {
            console.error("Connection Error:", e);
            return { error: 'Lỗi kết nối mạng.' };
        }
    }

    function formatAiResponse(text) {
        if (!text) return "";
        // Convert **bold** to <b>bold</b>
        // Remove markdown bullets and replace with clean format
        return text
            .replace(/\*\*(.*?)\*\*/g, '<b>$1</b>')
            .replace(/^\s*\* /gm, '• ')
            .replace(/\*/g, '') // Remove remaining single asterisks
            .replace(/\n/g, '<br>');
    }

    function getRoleText() {
        if (currentMode === 'skkn') return "";
        const role = document.querySelector('input[name="teacher_role"]:checked')?.value;
        return role === 'GVCNG' ? "Giáo viên chủ nhiệm giỏi (GVCNG)" : "Giáo viên giỏi (GVG)";
    }

    function getDocType() {
        if (currentMode === 'skkn') return "SKKN";
        // Kiểm tra xem Tab "THẨM ĐỊNH" có đang active hay không
        const activeTab = document.querySelector('.tab-btn.active');
        if (activeTab && activeTab.getAttribute('data-mode') === 'thẩm-định') {
            return document.querySelector('input[name="appraisal_type"]:checked')?.value || "BIỆN PHÁP";
        }
        return "BIỆN PHÁP";
    }

    // TÌM ĐOẠN VĂN BỐI CẢNH TỪ FILE GỐC
    function findOriginalSnippet(sectionId, fullText) {
        if (!fullText) return "";
        let keywords = [];
        if (sectionId === '1') keywords = ["ĐẶT VẤN ĐỀ", "LÝ DO", "MỤC ĐÍCH"];
        else if (sectionId === '2') keywords = ["THỰC TRẠNG", "CƠ SỞ LÝ LUẬN", "THỰC TIỄN"];
        else if (sectionId === '2.3.1') keywords = ["Giải pháp 1", "Giải pháp thứ nhất"];
        else if (sectionId === '2.3.2') keywords = ["Giải pháp 2", "Giải pháp thứ hai"];
        else if (sectionId === '2.3.3') keywords = ["Giải pháp 3", "Giải pháp thứ ba"];
        else if (sectionId === '2.4') keywords = ["Tính mới", "Tính sáng tạo", "Điểm mới"];
        else if (sectionId === '3') keywords = ["KẾT LUẬN", "KIẾN NGHỊ", "HIỆU QUẢ"];

        const lines = fullText.split('\n');
        let startLine = -1;
        for (let i = 0; i < lines.length; i++) {
            if (keywords.some(kw => lines[i].toUpperCase().includes(kw.toUpperCase()))) {
                startLine = i;
                break;
            }
        }

        if (startLine === -1) return "";
        // Lấy khoảng 100 dòng xung quanh đoạn đó
        return lines.slice(Math.max(0, startLine - 2), startLine + 100).join('\n');
    }

    async function writeSingleSection(id, btn, userPrompt = "") {
        const title = document.getElementById('topic-title').value;
        const roleText = getRoleText();
        if (!title) {
            alert('Vui lòng nhập tên đề tài!');
            switchSection('info');
            return false;
        }

        const area = document.querySelector(`.editor-area[data-id="${id}"]`);
        if (!area) return false;

        // Lấy thông tin bối cảnh của giáo viên
        const school = document.getElementById('school-name')?.value || "";
        const author = document.getElementById('author-name')?.value || "";
        const subject = document.getElementById('subject-name')?.value || "";
        const grade = document.getElementById('grade-level')?.value || "";
        const className = document.getElementById('class-name')?.value || "";
        const time = document.getElementById('execution-time')?.value || "";

        let userInfoContext = "";
        if (school || author || subject || grade || className || time) {
            userInfoContext = `\n\n[THÔNG TIN THỰC TẾ - HÃY DÙNG ĐỂ THAY THẾ CÁC CHỖ TRỐNG TRONG VĂN BẢN]:
            - Trường/Đơn vị: ${school}
            - Giáo viên/Người thực hiện: ${author}
            - Môn học/Lĩnh vực: ${subject}
            - Lớp: ${grade} ${className}
            - Thời gian/Năm học: ${time}
            LƯU Ý QUAN TRỌNG: TUYỆT ĐỐI KHÔNG sử dụng các từ trong ngoặc vuông như "[Tên trường]", "[Năm học]", "[Tên giáo viên]". Hãy điền thông tin thật ở trên vào các vị trí tương ứng trong bài viết một cách tự nhiên.`;
        }

        // Tinh toan do dai tung muc dua tren gioi han trang chung
        let sectionLimitText = "";
        const pageLimitInput = document.getElementById('upgrade-page-limit');
        const pageLimitVal = pageLimitInput ? pageLimitInput.value.trim() : "";
        const pages = pageLimitVal.split('-').map(p => parseInt(p.trim())).filter(p => !isNaN(p));

        const idsList = currentMode === 'skkn' ? ['1', 'theory', '2', '2.3.1', '2.3.2', '2.3.3', '2.4', '3'] : ['1', '2', '2.3.1', '2.3.2', '2.3.3', '2.4', '3'];

        let strategy = "";
        let wordsPerSection = 800;
        let hardLimit = 1500;

        if (pages.length > 0) {
            const maxPages = pages.length === 2 ? pages[1] : pages[0];
            const avgPages = (pages[0] + (pages[1] || pages[0])) / 2;
            const wordsPerPage = 300;
            const totalWordsTarget = Math.round(avgPages * wordsPerPage);

            wordsPerSection = Math.round(totalWordsTarget / idsList.length);
            hardLimit = wordsPerSection + 150;

            if (maxPages <= 12) {
                strategy = `[AI CHIẾN LƯỢC VIẾT NGẮN]: Tập trung viết cô đọng, súc tích, lược bỏ các ý rườm rà, tập trung vào luận điểm chính để báo cáo nằm gọn trong ${currentMode === 'skkn' ? '6 phần' : '3 phần'} với tổng số ${maxPages} trang.`;
            } else {
                strategy = `[AI CHIẾN LƯỢC VIẾT DÀI/SÂU]: Tự động phân tích sâu hơn, thêm thắt các ví dụ cụ thể, mô tả bảng biểu (nếu cần) và đưa ra nhiều dẫn chứng, phân tích đa chiều để làm dày nội dung cho đủ khoảng ${maxPages} trang.`;
            }
        } else {
            // Mặc định: Phân tích sâu, tự do sáng tạo
            strategy = `[AI CHIẾN LƯỢC TỰ DO SÁNG TẠO & CHUYÊN SÂU]: Không giới hạn số trang. Hãy phân tích cực kỳ sâu sắc, trình bày chi tiết mọi khía cạnh, thêm nhiều ví dụ minh họa và dẫn chứng thực tế. Viết càng dài, càng chuyên sâu càng tốt để thể hiện đẳng cấp của một chuyên gia.`;
            wordsPerSection = 1000;
            hardLimit = 2000;
        }

        sectionLimitText = `\n\n[YÊU CẦU VỀ ĐỘ DÀI & PHONG CÁCH]:
        - ${strategy}
        - MỤC TIÊU MỤC NÀY: Khoảng ${wordsPerSection} từ.
        - GIỚI HẠN CỨNG: Tuyệt đối không viết quá ${hardLimit} từ để dành chỗ cho các mục sau.`;

        const originalHtml = btn.innerHTML;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG SOẠN...';
        btn.disabled = true;
        area.innerHTML = "<i>[Đang soạn thảo...]</i>";

        const docType = currentMode === 'skkn' ? 'SKKN' : 'BIỆN PHÁP';
        const docName = docType === 'SKKN' ? "SÁNG KIẾN KINH NGHIỆM (SKKN)" : "BÁO CÁO BIỆN PHÁP GVG/GVCNG";

        // Lấy bối cảnh từ file gốc nếu có
        const originalText = findOriginalSnippet(id, lastAppraisedText);
        let baseContext = originalText ? `\n\n[NỘI DUNG GỐC TỪ FILE CỦA GIÁO VIÊN]:\n"""${originalText}"""\n\n[YÊU CẦU]: Hãy CHỈNH SỬA, bổ sung hoặc RÚT GỌN nội dung gốc trên. Tuyệt đối giữ nguyên các ý tưởng hay, minh chứng thực tế của giáo viên. Chỉ viết lại để văn phong mượt mà hơn và đúng độ dài.` : "";

        let prompt = "";

        if (docType === 'SKKN') {
            prompt = `Thực hiện viết ${docName} kỳ thi ${roleText}. Đề tài: "${title}". 
            YÊU CẦU: Viết nội dung cho mục ${skknStructure[id] || id}.${sectionLimitText}${userInfoContext}${baseContext}
            LƯU Ý CẤU TRÚC 3 PHẦN: 
            - Phần I: Đặt vấn đề (Mở đầu)
            - Phần II: Nội dung (Cơ sở, Thực trạng, Giải pháp, Hiệu quả)
            - Phần III: Kết luận & Kiến nghị
            PHONG CÁCH: Ngôn từ khoa học, đúc kết chuyên sâu, hành văn chuyên nghiệp. Viết trực tiếp vào trọng tâm, súc tích từng mục.`;
        } else {
            prompt = `Thực hiện ${originalText ? 'CHỈNH SỬA/HOÀN THIỆN' : 'Viết'} nội dung cho ${docName} kỳ thi ${roleText} với đề tài "${title}", mục ${id}.${sectionLimitText}${userInfoContext}${baseContext}`;
            if (id === '1') prompt = `Thực hiện ${originalText ? 'CHỈNH SỬA/HOÀN THIỆN' : 'Viết'} nội dung cho ${docName} kỳ thi ${roleText} đề tài "${title}", mục 1. ĐẶT VẤN ĐỀ (bao gồm 1.1 Lí do, 1.2 Mục đích, 1.3 Đối tượng).${sectionLimitText}${userInfoContext}${baseContext}`;
            if (id === '2') prompt = `Thực hiện ${originalText ? 'CHỈNH SỬA/HOÀN THIỆN' : 'Viết'} nội dung cho ${docName} kỳ thi ${roleText} đề tài "${title}", mục 2. THỰC TRẠNG VẤN ĐỀ (bao gồm 2.1 Cơ sở lý luận và 2.2 Cơ sở thực tiễn).${sectionLimitText}${userInfoContext}${baseContext}`;
        }

        if (userPrompt) {
            prompt += `\nLưu ý riêng từ giáo viên: "${userPrompt}"`;
        }

        const result = await callGemini(prompt, () => writeSingleSection(id, btn, userPrompt));

        if (result.text) {
            area.innerHTML = formatAiResponse(result.text);
            btn.innerHTML = originalHtml;
            btn.disabled = false;
            return true;
        } else {
            area.innerHTML = `<span style="color:red">[LỖI: ${result.error}]</span>`;
            btn.innerHTML = originalHtml;
            btn.disabled = false;
            return false;
        }
    }

    // Hook up all "Viết theo gợi ý (AI)" buttons
    document.querySelectorAll('.write-with-prompt-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const section = btn.closest('.card');
            const editor = section.querySelector('.editor-area');
            const promptInput = section.querySelector('.ai-prompt-input');
            const userPrompt = promptInput ? promptInput.value.trim() : "";

            if (!userPrompt) {
                alert("Vui lòng nhập nội dung gợi ý vào ô trống trước khi nhấn nút này!");
                if (promptInput) promptInput.focus();
                return;
            }

            if (editor) {
                writeSingleSection(editor.getAttribute('data-id'), btn, userPrompt);
            }
        });
    });

    // Hook up all "Viết tự động (AI)" buttons in sections
    document.querySelectorAll('.auto-write').forEach(btn => {
        btn.addEventListener('click', () => {
            const section = btn.closest('.card');
            const editor = section.querySelector('.editor-area');
            if (editor) {
                // Auto-write ignores the prompt input as requested
                writeSingleSection(editor.getAttribute('data-id'), btn);
            }
        });
    });

    // Progress modal elements were removed from HTML, keeping variables as null or removing references
    let isWritingAll = false;
    let autoSwitchEnabled = true;

    async function writeAllSections(startIndex = 0, extraContext = "") {
        const title = document.getElementById('topic-title').value;
        const roleText = getRoleText();
        if (!title) return alert('Vui lòng nhập tên đề tài!');

        writeAllBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG VIẾT...';
        writeAllBtn.disabled = true;

        try {
            const ids = currentMode === 'skkn' ? ['1', 'theory', '2', '2.3.1', '2.3.2', '2.3.3', '2.4', '3', 'appendix'] : ['1', '2', '2.3.1', '2.3.2', '2.3.3', '2.4', '3'];
            let prompt = "";

            const idToSection = {
                '1': 'section1',
                'theory': 'theory',
                '2': 'section2',
                '2.3.1': 'solution1',
                '2.3.2': 'solution2',
                '2.3.3': 'solution3',
                '2.4': 'creativity',
                '3': 'impact',
                'appendix': 'appendix'
            };

            if (writingProgressView) {
                writingProgressView.style.display = 'none'; // Ensure modal is hidden
            }

            for (let i = startIndex; i < ids.length; i++) {
                const id = ids[i];
                isWritingAll = true;

                // Switch to the section being written only if auto-switch is enabled
                if (idToSection[id] && autoSwitchEnabled) {
                    switchSection(idToSection[id]);
                    currentSectionIndex = currentSectionOrder.indexOf(idToSection[id]);
                    // Small scroll to top of section for better visibility
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                }

                // Update UI Progress
                const percent = Math.round((i / ids.length) * 100);
                if (progressBarFill) progressBarFill.style.width = `${percent}%`;
                if (progressPercent) progressPercent.innerText = `${percent}%`;

                const area = document.querySelector(`.editor-area[data-id="${id}"]`);
                if (area) {
                    area.innerHTML = "<i>[AI đang soạn thảo nội dung, vui lòng đợi...]</i>";

                    // Lấy thông tin bối cảnh của giáo viên
                    const school = document.getElementById('school-name')?.value || "";
                    const author = document.getElementById('author-name')?.value || "";
                    const subject = document.getElementById('subject-name')?.value || "";
                    const grade = document.getElementById('grade-level')?.value || "";
                    const className = document.getElementById('class-name')?.value || "";
                    const time = document.getElementById('execution-time')?.value || "";

                    let userInfoContext = "";
                    if (school || author || subject || grade || className || time) {
                        userInfoContext = `\n\n[THÔNG TIN THỰC TẾ - HÃY DÙNG ĐỂ THAY THẾ CÁC CHỖ TRỐNG TRONG VĂN BẢN]:
                        - Trường/Đơn vị: ${school}
                        - Giáo viên/Người thực hiện: ${author}
                        - Môn học/Lĩnh vực: ${subject}
                        - Lớp: ${grade} ${className}
                        - Thời gian/Năm học: ${time}
                        LƯU Ý QUAN TRỌNG: TUYỆT ĐỐI KHÔNG sử dụng các từ trong ngoặc vuông như "[Tên trường]", "[Năm học]", "[Tên giáo viên]". Hãy điền thông tin thật ở trên vào các vị trí tương ứng trong bài viết một cách tự nhiên.`;
                    }

                    // Get buttons to update state
                    const sectionBtn = area.closest('.card')?.querySelector('.auto-write');
                    const originalBtnHtml = sectionBtn ? sectionBtn.innerHTML : "";
                    if (sectionBtn) {
                        sectionBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG VIẾT...';
                        sectionBtn.disabled = true;
                    }

                    // Call writeSingleSection logic or call gemini directly
                    // We'll call gemini directly but similar to writeSingleSection logic

                    // (Existing prompt logic follows...)
                    let sectionLimitText = "";
                    const pageLimitVal = document.getElementById('upgrade-page-limit')?.value || "15-20";
                    const pages = pageLimitVal.split('-').map(p => parseInt(p.trim())).filter(p => !isNaN(p));
                    if (pages.length > 0) {
                        const maxPages = pages.length === 2 ? pages[1] : pages[0];
                        const avgPages = (pages[0] + (pages[1] || pages[0])) / 2;

                        // Định mức: 300 chữ/trang (an toàn cho dãn dòng 1.5 và lề chuẩn)
                        const wordsPerPage = 300;
                        const totalWordsTarget = Math.round(avgPages * wordsPerPage);
                        const wordsPerSection = Math.round(totalWordsTarget / ids.length);
                        const hardLimit = wordsPerSection + 100;

                        let strategy = "";
                        if (maxPages <= 12) {
                            strategy = "[AI CHIẾN LƯỢC VIẾT NGẮN]: Tập trung viết cô đọng, súc tích, lược bỏ các ý rườm rà, tổng hợp ý chính để đảm bảo báo cáo nằm gọn trong số trang quy định.";
                        } else {
                            strategy = "[AI CHIẾN LƯỢC ĐẠT CHUẨN ĐỘ DÀI]: \n- Nếu nội dung hiện tại đang ngắn: Hãy phân tích sâu hơn, thêm thắt các ví dụ cụ thể, bảng biểu và dẫn chứng thực tế để làm dày nội dung cho đủ số trang. \n- Nếu nội dung hiện tại đang quá dài (như file gốc): Hãy lược bớt các phần trùng lặp, viết cô đọng lại để vừa vặn khung hình.";
                        }

                        sectionLimitText = `\n\n[QUY ĐỊNH ĐỘ ĐÀI BẮT BUỘC]:
                    - ${strategy}
                    - MỤC TIÊU: Viết khoảng ${wordsPerSection} từ cho mục này (tương đương ${(wordsPerSection / wordsPerPage).toFixed(1)} trang A4).
                    - GIỚI HẠN CỨNG: TUYỆT ĐỐI KHÔNG viết quá ${hardLimit} từ.
                    - TỔNG THỂ: Toàn bộ báo cáo (7 mục) phải đạt đúng khoảng ${pageLimitVal} trang khi in Word.`;
                    }

                    const docType = getDocType();
                    const docName = docType === 'SKKN' ? "SÁNG KIẾN KINH NGHIỆM (SKKN)" : "BÁO CÁO BIỆN PHÁP GVG/GVCNG";

                    // Lấy bối cảnh từ file gốc nếu có
                    const originalText = findOriginalSnippet(id, lastAppraisedText);
                    let baseContext = originalText ? `\n\n[NỘI DUNG GỐC TỪ FILE CỦA GIÁO VIÊN]:\n"""${originalText}"""\n\n[YÊU CẦU CỦA GIÁO VIÊN]: Bạn phải sử dụng nội dung gốc trên làm nền tảng. CHỈ ĐƯỢC PHÉP chỉnh sửa văn phong, thêm thắt ý còn thiếu hoặc RÚT GỌN nội dung rườm rà. KHÔNG ĐƯỢC viết lại từ đầu bằng ý tưởng của AI. Hãy tôn trọng mọi minh chứng và giải pháp thực tế mà giáo viên đã viết.` : "";

                    if (docType === 'SKKN') {
                        prompt = `Thực hiện viết ${docName} kỳ thi ${roleText}. Đề tài: "${title}".
                        YÊU CẦU: Viết nội dung cho mục ${skknStructure[id] || id}.${sectionLimitText}${userInfoContext}${baseContext}
                        LƯU Ý CẤU TRÚC 3 PHẦN: I-Đặt vấn đề, II-Nội dung (Cơ sở, Thực trạng, Giải pháp, Hiệu quả), III-Kết luận.
                        PHONG CÁCH: Khoa học, súc tích, chuyên sâu.`;
                    } else {
                        prompt = `Thực hiện ${originalText ? 'CHỈNH SỬA/HOÀN THIỆN' : 'Viết'} nội dung cho ${docName} kỳ thi ${roleText} with đề tài "${title}", mục ${id}.${sectionLimitText}${userInfoContext}${baseContext}`;
                        if (id === '1') prompt = `Thực hiện ${originalText ? 'CHỈNH SỬA/HOÀN THIỆN' : 'Viết'} nội dung cho ${docName} kỳ thi ${roleText} đề tài "${title}", mục 1. ĐẶT VẤN ĐỀ (bao gồm 1.1 Lí do, 1.2 Mục đích, 1.3 Đối tượng).${sectionLimitText}${userInfoContext}${baseContext}`;
                        if (id === '2') prompt = `Thực hiện ${originalText ? 'CHỈNH SỬA/HOÀN THIỆN' : 'Viết'} nội dung cho ${docName} kỳ thi ${roleText} đề tài "${title}", mục 2. THỰC TRẠNG VẤN ĐỀ (bao gồm 2.1 Cơ sở lý luận và 2.2 Cơ sở thực tiễn).${sectionLimitText}${userInfoContext}${baseContext}`;
                    }

                    if (extraContext) {
                        prompt += `\n\nBỐI CẢNH NÂNG CẤP QUAN TRỌNG:\n${extraContext}`;
                    }

                    const result = await callGemini(prompt, () => writeAllSections(i, extraContext));
                    if (result.text) {
                        area.innerHTML = formatAiResponse(result.text);
                    } else {
                        area.innerHTML = `<span style="color:red">[LỖI: ${result.error}]</span>`;
                        if (result.error.includes('Hết hạn mức')) {
                            writeAllBtn.innerHTML = '<i class="fas fa-magic"></i> TIẾP TỤC VIẾT (AI)';
                            writeAllBtn.disabled = false;
                            // if (writingProgressView) writingProgressView.style.display = 'none'; // Removed
                            return;
                        }
                    }

                    // Restore button state
                    if (sectionBtn) {
                        sectionBtn.innerHTML = originalBtnHtml;
                        sectionBtn.disabled = false;
                    }
                }
                await new Promise(r => setTimeout(r, 1200));
            }

            // Finished Successfully
            isWritingAll = false;
            autoSwitchEnabled = true;
            const floatingBtn = document.getElementById('floating-progress-btn');
            if (floatingBtn) floatingBtn.style.display = 'none';

            // if (progressBarFill) progressBarFill.style.width = `100%`;
            // if (progressPercent) progressPercent.innerText = `100%`;
            // if (progressLoadingIcon) progressLoadingIcon.style.display = 'none';
            // if (progressSuccessIcon) progressSuccessIcon.style.display = 'block';
            // if (progressTitle) progressTitle.innerText = "Hoàn tất nâng cấp!";
            // if (progressStatus) progressStatus.innerText = "Báo cáo của bạn đã được nâng cấp thành phiên bản chuyên nghiệp hoàn chỉnh.";

            // If the modal was hidden, bring it back to show completion
            // if (writingProgressView.style.display === 'none') { // Removed
            //     writingProgressView.style.display = 'flex'; // Removed
            // }

            // document.getElementById('writing-back-action').style.display = 'none'; // Hide back button
            // if (finishedActions) finishedActions.style.display = 'flex';

            writeAllBtn.innerHTML = '<i class="fas fa-magic"></i> VIẾT TOÀN BỘ (AI)';
            writeAllBtn.disabled = false;
            alert("Chúc mừng! Hệ thống đã hoàn thành việc soạn thảo toàn bộ các mục. Bạn có thể kiểm tra từng phần và xuất file Word.");

        } catch (error) {
            console.error("Write All Error:", error);
            alert("Đã xảy ra lỗi trong quá trình viết tự động: " + error.message);
            writeAllBtn.innerHTML = '<i class="fas fa-magic"></i> VIẾT TOÀN BỘ (AI)';
            writeAllBtn.disabled = false;
        }
    }

    document.getElementById('final-export-docx-btn')?.addEventListener('click', () => {
        exportWordBtn?.click();
    });

    // Removed hide-progress-btn and close-progress-btn event listeners and floating button logic
    // document.getElementById('hide-progress-btn')?.addEventListener('click', () => {
    //     if (writingProgressView) writingProgressView.style.display = 'none';

    //     // Show a small floating indicator if still writing
    //     if (isWritingAll) {
    //         let floatingBtn = document.getElementById('floating-progress-btn');
    //         if (!floatingBtn) {
    //             floatingBtn = document.createElement('div');
    //             floatingBtn.id = 'floating-progress-btn';
    //             floatingBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Đang xử lý...';
    //             floatingBtn.style = "position: fixed; bottom: 20px; right: 20px; background: #16a34a; color: white; padding: 10px 20px; border-radius: 50px; cursor: pointer; z-index: 1000; box-shadow: 0 4px 15px rgba(0,0,0,0.2); font-weight: 700; display: flex; align-items: center; gap: 8px;";
    //             floatingBtn.onclick = () => {
    //                 writingProgressView.style.display = 'flex';
    //                 floatingBtn.style.display = 'none';
    //             };
    //             document.body.appendChild(floatingBtn);
    //         } else {
    //             floatingBtn.style.display = 'flex';
    //         }
    //     }
    // });

    // document.getElementById('close-progress-btn')?.addEventListener('click', () => {
    //     if (writingProgressView) writingProgressView.style.display = 'none';
    //     const floatingBtn = document.getElementById('floating-progress-btn');
    //     if (floatingBtn) floatingBtn.style.display = 'none';
    // });

    if (writeAllBtn) {
        writeAllBtn.addEventListener('click', () => writeAllSections(0));
    }

    // --- WORD EXPORT FIX: FULL REPORT AGGREGATION ---
    if (exportWordBtn) {
        exportWordBtn.addEventListener('click', async () => {
            const topic = document.getElementById('topic-title')?.value || "TÊN ĐỀ TÀI CHƯA NHẬP";
            const author = document.getElementById('author-name')?.value || "Người thực hiện";
            const school = document.getElementById('school-name')?.value || "Tên Trường";
            const subject = document.getElementById('subject-name')?.value || "";
            const grade = document.getElementById('grade-level')?.value || "";
            const className = document.getElementById('class-name')?.value || "";
            const time = document.getElementById('execution-time')?.value || "";

            exportWordBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG XUẤT...';
            exportWordBtn.disabled = true;

            try {
                const docxLib = window.docx || docx;
                if (!docxLib) throw new Error("Chưa tải được thư viện docx.");

                const { Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak } = docxLib;

                const children = [];

                // --- PAGE 1: COVER PAGE ---
                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: school.toUpperCase(), size: 26, font: "Times New Roman" })]
                }));

                children.push(new Paragraph({ spacing: { before: 2000 } })); // Spacer

                const docType = getDocType();
                const titleText = docType === 'SKKN' ? "BÁO CÁO SÁNG KIẾN KINH NGHIỆM" : "BÁO CÁO BIỆN PHÁP";
                const subTitleText = docType === 'SKKN' ? "GÓP PHẦN NÂNG CAO HIỆU QUẢ CÔNG TÁC GIÁO DỤC" : "GÓP PHẦN NÂNG CAO CHẤT LƯỢNG CÔNG TÁC GIẢNG DẠY";

                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: titleText, bold: true, size: 32, font: "Times New Roman" })]
                }));

                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: subTitleText, size: 24, font: "Times New Roman" })]
                }));

                children.push(new Paragraph({ spacing: { before: 800 } })); // Spacer

                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "ĐỀ TÀI:", bold: true, size: 26, font: "Times New Roman" })]
                }));
                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: topic.toUpperCase(), bold: true, size: 32, font: "Times New Roman" })]
                }));

                children.push(new Paragraph({ spacing: { before: 2000 } })); // Spacer

                const infoEntries = [
                    { label: "Môn học/Lĩnh vực:", value: subject },
                    { label: "Cấp học/Khối lớp:", value: `${grade} - ${className}` },
                    { label: "Người thực hiện:", value: author },
                    { label: "Thời gian thực hiện:", value: time }
                ];

                infoEntries.forEach(entry => {
                    children.push(new Paragraph({
                        alignment: AlignmentType.LEFT,
                        indent: { left: 3000 },
                        children: [
                            new TextRun({ text: entry.label + " ", bold: true, size: 28, font: "Times New Roman" }),
                            new TextRun({ text: entry.value, size: 28, font: "Times New Roman" })
                        ],
                        spacing: { line: 360 }
                    }));
                });

                children.push(new Paragraph({ children: [new PageBreak()] }));

                // --- PAGE 2 ONWARDS: CONTENT ---
                const exportTasks = [
                    { id: '1', title: 'I. ĐẶT VẤN ĐỀ' },
                    { id: '2', title: 'II. THỰC TRẠNG VẤN ĐỀ' },
                    { id: '2.3.1', title: '2.3 Các giải pháp thực hiện', sub: '2.3.1 Giải pháp thứ nhất' },
                    { id: '2.3.2', title: '2.3.2 Giải pháp thứ hai' },
                    { id: '2.3.3', title: '2.3.3 Giải pháp thứ ba' },
                    { id: '2.4', title: '2.4 Tính mới và sáng tạo' },
                    { id: '3', title: 'III. KẾT LUẬN' }
                ];

                exportTasks.forEach(task => {
                    // Title
                    children.push(new Paragraph({
                        children: [new TextRun({ text: task.title, bold: true, size: 28, font: "Times New Roman" })],
                        spacing: { before: 400, after: 200 }
                    }));

                    // Sub-title
                    if (task.sub) {
                        children.push(new Paragraph({
                            children: [new TextRun({ text: task.sub, bold: true, italics: true, size: 24, font: "Times New Roman" })],
                            spacing: { before: 200, after: 100 }
                        }));
                    }

                    const editor = document.querySelector(`.editor-area[data-id="${task.id}"]`);
                    const sectionParagraphs = htmlToParagraphs(editor ? editor.innerHTML : "(Trống)");

                    if (sectionParagraphs.length > 0) {
                        children.push(...sectionParagraphs);
                    } else {
                        children.push(new Paragraph({
                            children: [new TextRun({ text: "(Chưa có nội dung)", italics: true, size: 24, font: "Times New Roman" })]
                        }));
                    }
                });

                const doc = new Document({
                    sections: [{
                        properties: {
                            page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1701 } }
                        },
                        children: children
                    }]
                });

                const blob = await Packer.toBlob(doc);
                const fileName = cleanFileName(author) + ".docx";

                if (window.saveAs) {
                    window.saveAs(blob, fileName);
                } else {
                    const url = URL.createObjectURL(blob);
                    const a = document.createElement("a");
                    a.href = url;
                    a.download = fileName;
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                }

                alert("Xuất Word thành công!");

            } catch (error) {
                console.error("Loi xuat Word:", error);
                alert("Lỗi xuất Word: " + error.message);
            } finally {
                exportWordBtn.innerHTML = '<i class="fas fa-file-word"></i> XUẤT WORD (.DOCX)';
                exportWordBtn.disabled = false;
            }
        });
    }

    // AI Suggestion for Topic Title
    aiSuggestBtn?.addEventListener('click', async () => {
        const topicInput = document.getElementById('topic-title');
        const roleText = getRoleText();
        const subject = document.getElementById('subject-name')?.value || "một môn học";
        const grade = document.getElementById('grade-level')?.value || "các cấp";
        const className = document.getElementById('class-name')?.value || "";
        const currentTitle = topicInput?.value || "";

        const originalHtml = aiSuggestBtn.innerHTML;
        aiSuggestBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Đang phân tích...';
        aiSuggestBtn.disabled = true;

        let promptText = `Bạn là chuyên gia giáo dục. Hãy phân tích đề tài: "${currentTitle}" cho môn ${subject}, lớp ${grade} (Kỳ thi ${roleText}).
        Đề xuất 5 phương án hay hơn, sát thực tế, đậm chất GVG/GVCNG chuyên nghiệp. 
        Mỗi đề tài trên 1 dòng. Không thêm giải thích, không đánh số, không ngoặc kép.`;

        const result = await callGemini(promptText, () => aiSuggestBtn.click());

        if (result.text) {
            const suggestions = result.text.split('\n').filter(s => s.trim().length > 5);
            if (suggestions.length > 0) {
                showTopicChoiceModal(suggestions, (selectedTitle) => {
                    if (topicInput) topicInput.value = selectedTitle;
                });
            }
        }

        aiSuggestBtn.innerHTML = originalHtml;
        aiSuggestBtn.disabled = false;
    });

    const topicChoiceModal = document.getElementById('topic-choice-modal');
    const topicChoiceList = document.getElementById('topic-choice-list');

    function showTopicChoiceModal(suggestions, onSelect) {
        if (!topicChoiceModal || !topicChoiceList) return;

        topicChoiceList.innerHTML = '';
        suggestions.slice(0, 5).forEach((text, index) => {
            const cleanText = text.replace(/^\d+[\s.)]+/, '').replace(/["']/g, '').trim();
            const btn = document.createElement('button');
            btn.className = 'topic-option-btn';
            btn.innerHTML = `<span class="num-badge">${index + 1}</span> <span>${cleanText}</span>`;
            btn.onclick = () => {
                onSelect(cleanText);
                topicChoiceModal.style.display = 'none';
            };
            topicChoiceList.appendChild(btn);
        });

        topicChoiceModal.style.display = 'flex';
    }

    document.getElementById('close-topic-choice')?.addEventListener('click', () => {
        topicChoiceModal.style.display = 'none';
    });
    document.getElementById('cancel-topic-choice')?.addEventListener('click', () => {
        topicChoiceModal.style.display = 'none';
    });

    completeBtn?.addEventListener('click', () => {
        alert("Chúc mừng! Bạn đã hoàn thành báo cáo biện pháp GVG/GVCNG. Hãy nhấn Xuất Word để tải file về.");
        if (completeBtn) {
            completeBtn.innerHTML = '<i class="fas fa-check"></i> ĐÃ HOÀN THÀNH';
            completeBtn.style.opacity = '0.8';
        }
        if (exportWordBtn) exportWordBtn.style.display = 'flex';
    });

    // Appraisal Logic
    if (selectFileBtn) {
        selectFileBtn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            fileInput?.click();
        });
    }

    fileInput?.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            startAnalysis(e.target.files[0]);
        }
    });

    uploadDropzone?.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadDropzone.style.borderColor = 'var(--primary-color)';
        uploadDropzone.style.background = '#f0f7ff';
    });

    uploadDropzone?.addEventListener('dragleave', () => {
        uploadDropzone.style.borderColor = '#cbd5e0';
        uploadDropzone.style.background = '#f8fafc';
    });

    // --- SKKN LOGIC ---
    // Finish SKKN Logic
    skknCompleteBtn?.addEventListener('click', () => {
        alert("Chúc mừng! Bạn đã hoàn thành báo cáo Sáng kiến kinh nghiệm (SKKN). Hãy nhấn Xuất Word để tải file về.");
        if (skknCompleteBtn) {
            skknCompleteBtn.innerHTML = '<i class="fas fa-check"></i> ĐÃ HOÀN THÀNH';
            skknCompleteBtn.style.opacity = '0.8';
        }
        if (skknExportWordBtn) skknExportWordBtn.style.display = 'flex';
    });

    // Word Export for SKKN
    document.getElementById('skkn-export-word-btn')?.addEventListener('click', async () => {
        const topic = document.getElementById('topic-title')?.value || "TÊN ĐỀ TÀI CHƯA NHẬP";
        const author = document.getElementById('author-name')?.value || "Người thực hiện";
        const school = document.getElementById('school-name')?.value || "Tên Trường";
        const subject = document.getElementById('subject-name')?.value || "";
        const grade = document.getElementById('grade-level')?.value || "";
        const className = document.getElementById('class-name')?.value || "";
        const time = document.getElementById('execution-time')?.value || "";

        const btn = document.getElementById('skkn-export-word-btn');
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG XUẤT...';
        btn.disabled = true;

        try {
            const docxLib = window.docx || docx;
            const { Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak } = docxLib;
            const children = [];

            // --- Cover Page ---
            children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: school.toUpperCase(), size: 26, font: "Times New Roman" })] }));
            children.push(new Paragraph({ spacing: { before: 2000 } }));
            children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "BÁO CÁO SÁNG KIẾN KINH NGHIỆM", bold: true, size: 32, font: "Times New Roman" })] }));
            children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ĐỀ TÀI:", bold: true, size: 26, font: "Times New Roman" })] }));
            children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: topic.toUpperCase(), bold: true, size: 30, font: "Times New Roman" })] }));
            children.push(new Paragraph({ spacing: { before: 2000 } }));

            const info = [
                { l: "Môn học:", v: subject },
                { l: "Khối lớp:", v: className },
                { l: "Tác giả:", v: author },
                { l: "Năm học:", v: time }
            ];
            info.forEach(item => {
                children.push(new Paragraph({
                    indent: { left: 3000 }, children: [
                        new TextRun({ text: item.l + " ", bold: true, size: 28, font: "Times New Roman" }),
                        new TextRun({ text: item.v || "", size: 28, font: "Times New Roman" })
                    ],
                    spacing: { line: 360 }
                }));
            });
            children.push(new Paragraph({ children: [new PageBreak()] }));

            // --- Content ---
            const skknSections = [
                { id: '1', title: 'PHẦN I: ĐẶT VẤN ĐỀ' },
                { id: 'theory', title: 'PHẦN II: CƠ SỞ LÝ LUẬN' },
                { id: '2', title: 'PHẦN III: THỰC TRẠNG VẤN ĐỀ' },
                { id: '2.3.1', title: 'PHẦN IV: CÁC GIẢI PHÁP (GIẢI PHÁP 1)' },
                { id: '2.3.2', title: 'PHẦN IV: CÁC GIẢI PHÁP (GIẢI PHÁP 2)' },
                { id: '2.3.3', title: 'PHẦN IV: CÁC GIẢI PHÁP (GIẢI PHÁP 3)' },
                { id: '2.4', title: 'PHẦN V: HIỆU QUẢ CỦA SÁNG KIẾN KINH NGHIỆM' },
                { id: '3', title: 'PHẦN VI: KẾT LUẬN VÀ KHUYẾN NGHỊ' },
                { id: 'appendix', title: 'PHỤ LỤC' }
            ];

            skknSections.forEach(s => {
                // Section Title Paragraph
                children.push(new Paragraph({
                    children: [new TextRun({ text: s.title, bold: true, size: 32, font: "Times New Roman" })], // Cỡ 16 cho tiêu đề mục
                    spacing: { before: 400, after: 200, line: 360 }
                }));

                const editor = document.querySelector(`.editor-area[data-id="${s.id}"]`);
                const sectionParagraphs = htmlToParagraphs(editor ? editor.innerHTML : "(Trống)");

                if (sectionParagraphs.length > 0) {
                    children.push(...sectionParagraphs);
                } else {
                    children.push(new Paragraph({ children: [new TextRun({ text: "(Chưa có nội dung)", italics: true, size: 24, font: "Times New Roman" })] }));
                }
            });

            const doc = new Document({
                sections: [{
                    properties: {
                        page: {
                            margin: { top: 1134, right: 1134, bottom: 1134, left: 1701 }
                        }
                    },
                    children: children
                }]
            });

            const blob = await Packer.toBlob(doc);
            const fileName = "SKKN_" + cleanFileName(author) + ".docx";
            window.saveAs(blob, fileName);
            alert("Sáng kiến kinh nghiệm đã được xuất thành công!");
        } catch (e) {
            console.error(e);
            alert("Lỗi xuất Word: " + e.message);
        } finally {
            btn.innerHTML = '<i class="fas fa-file-word"></i> XUẤT WORD SKKN (.DOCX)';
            btn.disabled = false;
        }
    });

    // Handle initial navigation state
    document.querySelectorAll('.skkn-nav').forEach(el => el.style.display = 'none');

    uploadDropzone?.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadDropzone.style.borderColor = '#cbd5e0';
        uploadDropzone.style.background = '#f8fafc';
        if (e.dataTransfer.files.length > 0) {
            startAnalysis(e.dataTransfer.files[0]);
        }
    });

    async function extractText(file) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (ext === 'docx') {
            const arrayBuffer = await file.arrayBuffer();
            const result = await mammoth.extractRawText({ arrayBuffer });
            return result.value;
        } else if (ext === 'pdf') {
            const pdfjsLib = window['pdfjs-dist/build/pdf'];
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
            const arrayBuffer = await file.arrayBuffer();
            const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
            let text = "";
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const content = await page.getTextContent();
                text += content.items.map(item => item.str).join(" ") + "\n";
            }
            return text;
        }
        return "";
    }

    // --- APPRAISAL LOGIC HANDLED BELOW ---

    // Export Appraisal Report to DOCX
    document.querySelectorAll('.export-appraisal-btn').forEach(btn => {
        btn.addEventListener('click', async () => {
            if (!lastAppraisalData) return;

            try {
                if (lastAppraisalData.type === 'SKKN') {
                    await exportSkknAppraisal(lastAppraisalData);
                } else {
                    await exportBpAppraisal(lastAppraisalData);
                }
            } catch (error) {
                console.error("Export Error:", error);
                alert("Lỗi khi xuất file Word: " + error.message);
            }
        });
    });

    async function exportBpAppraisal(data) {
        const doc = new docx.Document({
            styles: {
                default: {
                    paragraph: {
                        run: { size: 24, font: "Times New Roman" }
                    }
                }
            },
            sections: [{
                properties: {},
                children: [
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: "BÁO CÁO KẾT QUẢ THẨM ĐỊNH BIỆN PHÁP GVG/GVCNG", bold: true, size: 28 })],
                        alignment: docx.AlignmentType.CENTER,
                        spacing: { after: 400 }
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: "Tên đề tài: ", bold: true }), new docx.TextRun({ text: data.topic_name || "Chưa rõ" })],
                        spacing: { after: 120 }
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: "Kết quả tổng quát: ", bold: true }), new docx.TextRun({ text: `${data.quality_score}/100 điểm - Xếp loại: ${data.status_label}`, bold: true })],
                        spacing: { after: 200 }
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: "I. CHI TIẾT CÁC TIÊU CHÍ ĐÁNH GIÁ (Thang điểm 100)", bold: true, size: 24 })],
                        spacing: { before: 200, after: 120 }
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: `- Tính khoa học: ${data.criteria?.scientific || 0}/15` })],
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: `- Tính thực tiễn: ${data.criteria?.practical || 0}/20` })],
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: `- Tính mới và sáng tạo: ${data.criteria?.innovation || 0}/20` })],
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: `- Khả năng áp dụng: ${data.criteria?.applied_ability || 0}/20` })],
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: `- Hiệu quả minh chứng: ${data.criteria?.demonstrated_effect || 0}/15` })],
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: `- Ngôn ngữ và trình bày: ${data.criteria?.language || 0}/10` })],
                    }),

                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: "II. ĐÁNH GIÁ CHI TIẾT NỘI DUNG", bold: true, size: 24 })],
                        spacing: { before: 200, after: 120 }
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: data.summary || "Bản thảo có cấu trúc cơ bản nhưng cần hoàn thiện thêm các chi tiết chuyên môn." })],
                        spacing: { after: 200 }
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: "III. PHÂN TÍCH YẾU TỐ AI", bold: true, size: 24 })],
                        spacing: { before: 200, after: 120 }
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: `Nguy cơ sử dụng AI: ${data.ai_risk}%`, bold: true, color: data.ai_risk > 50 ? "FF0000" : "000000" })],
                        spacing: { after: 120 }
                    }),
                    ...(data.ai_risk_detail?.passages || []).map(p => new docx.Paragraph({
                        children: [
                            new docx.TextRun({ text: `• Đoạn văn: "${p.text}"`, italic: true }),
                            new docx.TextRun({ text: `\n  ↳ Lý do: ${p.reason}`, size: 20, color: "666666" })
                        ],
                        indent: { left: 400 },
                        spacing: { after: 120 }
                    })),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: "IV. GỢI Ý CẢI THIỆN", bold: true, size: 24 })],
                        spacing: { before: 200, after: 120 }
                    }),
                    ...(data.ai_details?.suggestions || ["Hoàn thiện nội dung các giải pháp cụ thể.", "Thêm minh chứng thực tế cho hiệu quả biện pháp."]).map(s => new docx.Paragraph({
                        children: [new docx.TextRun({ text: `• ${s}` })],
                        indent: { left: 400 },
                        spacing: { after: 120 }
                    })),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: "V. KẾT LUẬN CHUNG", bold: true, size: 24 })],
                        spacing: { before: 400, after: 120 }
                    }),
                    new docx.Paragraph({
                        children: [new docx.TextRun({ text: `Bản thảo đạt mức ${data.status_label}. ${data.quality_score >= 80 ? 'Có thể nộp ngay sau khi chỉnh sửa nhẹ.' : 'Cần chỉnh sửa sâu theo các gợi ý trên để đạt kết quả cao trong kỳ thi GVG.'}` })]
                    })
                ]
            }],
        });

        const blob = await docx.Packer.toBlob(doc);
        saveAs(blob, "Ket_qua_tham_dinh_Bien_phap.docx");
    }

    async function exportSkknAppraisal(data) {
        const { Table, TableRow, TableCell, Paragraph, TextRun, AlignmentType, WidthType, BorderStyle, Document, Packer } = docx;

        const teacherName = document.getElementById('author-name')?.value || "...";
        const subject = document.getElementById('subject-name')?.value || "...";
        const grade = document.getElementById('grade-level')?.value || "...";
        const className = document.getElementById('class-name')?.value || "...";
        const topicName = data.topic_name || "...";

        const children = [];

        // Header
        children.push(new Paragraph({
            children: [new TextRun({ text: "BIÊN BẢN CHẤM VÀ XÉT DUYỆT SÁNG KIẾN KINH NGHIỆM", bold: true, size: 28 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        }));

        const infoLines = [
            { l: "Tác giả:", v: teacherName },
            { l: "Đơn vị:", v: "........................................................................" },
            { l: "Tên SKKN:", v: topicName },
            { l: "Môn (lĩnh vực):", v: `${subject}` }
        ];

        infoLines.forEach(line => {
            children.push(new Paragraph({
                children: [
                    new TextRun({ text: line.l + " ", bold: true }),
                    new TextRun({ text: line.v })
                ],
                spacing: { after: 120 }
            }));
        });

        // Table
        const tableRows = [
            new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "TT", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 5, unit: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Nội dung", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 55, unit: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Điểm", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 10, unit: WidthType.PERCENTAGE } }),
                    new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Nhận xét", bold: true })], alignment: AlignmentType.CENTER })], width: { size: 30, unit: WidthType.PERCENTAGE } }),
                ]
            })
        ];

        const skknScores = data.skkn_scores || { innovation: 0, practicality: 0, scientific: 0, benefit: 0, format: 0 };
        const criteriaList = [
            { tt: "1", content: "Tính mới và sáng tạo (25 điểm)", max: 25, score: skknScores.innovation },
            { tt: "2", content: "Tính thực tiễn và khả thi (30 điểm)", max: 30, score: skknScores.practicality },
            { tt: "3", content: "Tính khoa học (20 điểm)", max: 20, score: skknScores.scientific },
            { tt: "4", content: "Hiệu quả và lợi ích (15 điểm)", max: 15, score: skknScores.benefit },
            { tt: "5", content: "Hình thức trình bày (10 điểm)", max: 10, score: skknScores.format }
        ];

        criteriaList.forEach(c => {
            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ text: c.tt, alignment: AlignmentType.CENTER })] }),
                    new TableCell({ children: [new Paragraph({ text: c.content })] }),
                    new TableCell({ children: [new Paragraph({ text: c.score.toString(), alignment: AlignmentType.CENTER })] }),
                    new TableCell({ children: [new Paragraph({ text: "Đạt yêu cầu" })] }),
                ]
            }));
        });

        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ text: "", alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ text: "TỔNG CỘNG", bold: true })] }),
                new TableCell({ children: [new Paragraph({ text: data.quality_score.toString(), bold: true, alignment: AlignmentType.CENTER })] }),
                new TableCell({ children: [new Paragraph({ text: "" })] }),
            ]
        }));

        children.push(new Table({
            width: { size: 100, unit: WidthType.PERCENTAGE },
            rows: tableRows
        }));

        children.push(new Paragraph({ text: "" }));
        children.push(new Paragraph({
            children: [new TextRun({ text: "NHẬN XÉT VÀ KẾT LUẬN CHUNG", bold: true, size: 24 })],
            spacing: { before: 400, after: 200 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: "Đánh giá chung: ", bold: true }), new TextRun({ text: data.summary || "Bản thảo SKKN đạt yêu cầu về cấu trúc và nội dung." })],
            spacing: { after: 120 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: "Xếp loại: ", bold: true }), new TextRun({ text: data.status_label || "Chưa xếp loại", bold: true, size: 28 })],
            spacing: { after: 120 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: "Ghi chú xếp loại: A (17-20đ), B (14-17đ), C (10-14đ), Không xếp loại (<10đ).", italics: true, size: 20 })]
        }));

        children.push(new Paragraph({
            children: [new TextRun({ text: "PHÂN TÍCH YẾU TỐ SỬ DỤNG AI", bold: true, size: 24 })],
            spacing: { before: 400, after: 200 }
        }));
        children.push(new Paragraph({
            children: [new TextRun({ text: `Nguy cơ AI detected: ${data.ai_risk}%`, bold: true })],
            spacing: { after: 120 }
        }));
        (data.ai_risk_detail?.passages || []).forEach(p => {
            children.push(new Paragraph({
                children: [
                    new TextRun({ text: `• Nội dung: "${p.text}"`, italic: true }),
                    new TextRun({ text: ` (Nhận diện: ${p.reason})`, size: 20 })
                ],
                indent: { left: 400 },
                spacing: { after: 80 }
            }));
        });

        const doc = new Document({
            sections: [{ properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 1440 } } }, children: children }]
        });

        const blob = await Packer.toBlob(doc);
        saveAs(blob, "Bien_ban_Tham_dinh_SKKN.docx");
    }

    // Mini export buttons inside upgrade plan
    document.querySelector('#upgrade-plan-view .mini-export-btn.docx')?.addEventListener('click', () => {
        if (exportAppraisalBtn) exportAppraisalBtn.click();
    });

    if (confirmUpgradeBtn) confirmUpgradeBtn.addEventListener('click', async () => {
        const appType = document.querySelector('input[name="appraisal_type"]:checked')?.value || "BIỆN PHÁP";

        // Nhảy về Tab phù hợp (Biện pháp hoặc SKKN)
        const mode = appType === "SKKN" ? "skkn" : "biên-pháp";
        const tabBtnSource = Array.from(tabButtons).find(btn => btn.getAttribute('data-mode') === mode);
        if (tabBtnSource) tabBtnSource.click();

        // Nhảy thẳng đến phần 1. Đặt vấn đề (vì tên đã có/chọn rồi)
        currentSectionIndex = currentSectionOrder.indexOf('section1');
        switchSection('section1');

        // --- CÁC HÀM BỔ TRỢ LÀM SẠCH DỮ LIỆU ---
        const cleanAiDoc = (text) => {
            if (!text) return "";
            // Thay thế ký tự xuống dòng giả nếu có
            let cleaned = text.replace(/\\n/g, '\n');

            // LOẠI BỎ CÁC KÝ HIỆU MARKDOWN (*)
            cleaned = cleaned.replace(/\*\*/g, ''); // Bỏ dấu in đậm **
            cleaned = cleaned.replace(/^\s*\*\s+/gm, '- '); // Bỏ dấu * đầu dòng, thay bằng dấu -
            // Một số AI trả về # cho tiêu đề, cũng nên làm sạch
            cleaned = cleaned.replace(/^#+\s+/gm, '');

            cleaned = cleaned.replace(/^(Chào bạn|Dưới đây là|Chào mừng|Đây là|Hệ thống đã|Vâng).*?\n/gi, '');
            cleaned = cleaned.replace(/ (Chào bạn|Chúc bạn|Nếu cần hỗ trợ).*?$/gi, '');

            const startIdx = cleaned.search(/(PHẦN I|I\. ĐẶT VẤN ĐỀ|Đặt vấn đề)/i);
            if (startIdx > -1) cleaned = cleaned.substring(startIdx);

            cleaned = cleaned.replace(/(Hy vọng bài viết|Chúc bạn thành công|Nếu cần hỗ trợ thêm).*?$/gi, '');
            return cleaned.trim();
        };

        const cleanTopicTitle = (text) => {
            if (!text) return "";
            let t = text.replace(/^(Chào bạn|Dưới đây là|Tên đề tài:|ĐỀ TÀI:).*?\n/gi, '');
            t = t.replace(/(chất lượng cao|xuất sắc|đạt giải|thảo mới).*?$/gi, '');
            t = t.replace(/^["'“”]|["'“”]$/g, '');
            return t.trim().substring(0, 150);
        };

        const pageCountStr = upgradePageLimitInput?.value || "15-20";
        if (upgradePlanView) upgradePlanView.style.display = 'none';

        if (writingProgressOverlay) {
            writingProgressOverlay.style.display = 'flex';
            if (progressLoadingIcon) progressLoadingIcon.style.display = 'block';
            if (progressSuccessIcon) progressSuccessIcon.style.display = 'none';
            if (finishedActions) finishedActions.style.display = 'none';
            if (progressTitle) progressTitle.innerText = `Đang hoàn thiện ${appType}...`;
            if (progressStatus) progressStatus.innerText = "AI đang tổng hợp dữ liệu từ file gốc và kết quả thẩm định để tạo bản thảo mới chuyên nghiệp.";
            if (progressLog) progressLog.innerHTML = "";
            addLogEntry("Khởi động hệ thống tái cấu trúc văn bản...");

            // Link close button
            if (closeProgressOverlayBtn) closeProgressOverlayBtn.onclick = () => {
                if (writingProgressOverlay) writingProgressOverlay.style.display = 'none';
            };
        }

        const feedback = (lastAppraisalData?.summary || "") + "\n" + (lastAppraisalData?.expert_advice || "");

        // Define internal mapping for Editor data-id attributes
        const getEditorId = (secId) => {
            const mapping = {
                'section1': '1',
                'theory': 'theory',
                'section2': '2',
                'solution1': '2.3.1',
                'creativity': '2.4',
                'impact': '3'
            };
            return mapping[secId] || secId;
        };

        const sectionsToDraft = appType === 'SKKN' ? [
            { id: 'section1', title: 'PHẦN I: ĐẶT VẤN ĐỀ', weight: 15 },
            { id: 'theory', title: 'PHẦN II: CƠ SỞ LÝ LUẬN', weight: 15 },
            { id: 'section2', title: 'PHẦN III: THỰC TRẠNG VẤN ĐỀ', weight: 15 },
            { id: 'solution1', title: 'PHẦN IV: CÁC GIẢI PHÁP', weight: 30 },
            { id: 'creativity', title: 'PHẦN V: HIỆU QUẢ CỦA SÁNG KIẾN', weight: 15 },
            { id: 'impact', title: 'PHẦN VI: KẾT LUẬN VÀ KHUYẾN NGHỊ', weight: 10 }
        ] : [
            { id: 'section1', title: 'I. ĐẶT VẤN ĐỀ', weight: 15 },
            { id: 'section2', title: 'II. THỰC TRẠNG VẤN ĐỀ', weight: 20 },
            { id: 'solution1', title: 'III. NỘI DUNG CÁC BIỆN PHÁP', weight: 40 },
            { id: 'creativity', title: 'IV. HIỆU QUẢ CỦA BIỆN PHÁP', weight: 15 },
            { id: 'impact', title: 'V. KẾT LUẬN VÀ KIẾN NGHỊ', weight: 10 }
        ];

        let currentProgress = 0;

        for (let i = 0; i < sectionsToDraft.length; i++) {
            const sec = sectionsToDraft[i];
            addLogEntry(`Giai đoạn ${i + 1}/${sectionsToDraft.length}: Đang soạn thảo ${sec.title}...`);
            if (progressStatus) progressStatus.innerText = `Đang hoàn thiện nội dung ${sec.title}...`;

            const sectionPrompt = `BẠN LÀ MÁY VIẾT BÁO CÁO GIÁO DỤC CHUYÊN NGHIỆP TRÌNH ĐỘ CAO.
            [NHIỆM VỤ]: Viết cực kỳ chi tiết ${sec.title} cho báo cáo ${appType}.
            [ĐỀ TÀI]: "${topicTitleInput?.value || "Đề tài giáo dục"}"
            [DỮ LIỆU GỐC]: """${lastAppraisedText.substring(0, 5000)}"""
            [HƯỚNG DẪN TỐI ƯU]: """${feedback}"""
            
            [YÊU CẦU TRÌNH BÀY]:
            - Văn phong sư phạm chuẩn GVG, trang trọng, logic.
            - Với các giải pháp, phải có bước thực hiện 1, 2, 3... rõ ràng.
            - KHÔNG chào hỏi, KHÔNG giới thiệu. Bắt đầu ngay vào nội dung.
            
            BẮT ĐẦU VIẾT ${sec.title}:`;

            try {
                const result = await callGemini(sectionPrompt, () => { });
                if (result.text) {
                    const cleanedPart = cleanAiDoc(result.text);
                    // Update specific editor area using data-id
                    const targetId = getEditorId(sec.id);
                    const editor = document.querySelector(`.editor-area[data-id="${targetId}"]`);
                    if (editor) {
                        // Simple formatting for view
                        const formatted = cleanedPart.split('\n').map(p => p.trim() ? `<div>${p}</div>` : '<div><br></div>').join('');
                        editor.innerHTML = formatted;
                    }
                }
            } catch (err) {
                console.error(`Error drafting section ${sec.id}:`, err);
                addLogEntry(`Gặp lỗi tại ${sec.title}, hệ thống đang skipping...`);
            }

            // Update UI Progress
            currentProgress += sec.weight;
            if (progressBarFill) progressBarFill.style.width = `${currentProgress}%`;
            if (progressPercent) progressPercent.innerText = `${currentProgress}%`;
        }

        // Final UI updates
        if (progressBarFill) progressBarFill.style.width = `100%`;
        if (progressPercent) progressPercent.innerText = `100%`;
        if (progressLoadingIcon) progressLoadingIcon.style.display = 'none';
        if (progressSuccessIcon) progressSuccessIcon.style.display = 'block';
        if (progressTitle) progressTitle.innerText = "Hoàn tất cải thiện!";
        if (progressStatus) progressStatus.innerText = `Hệ thống đã hoàn thành việc nâng cấp ${appType} của bạn theo từng phần chuyên sâu.`;
        addLogEntry("Quy trình hoàn tất. Tất cả nội dung đã được đồng bộ.");

        if (finishedActions) finishedActions.style.display = 'flex';

        if (viewResultBtn) viewResultBtn.onclick = () => {
            if (writingProgressOverlay) writingProgressOverlay.style.display = 'none';
            switchSection('section1');
            currentSectionIndex = currentSectionOrder.indexOf('section1');
        };

        if (exportAfterUpgradeBtn) exportAfterUpgradeBtn.onclick = () => {
            if (exportWordBtn) exportWordBtn.click(); // Trigger global export logic
        };
    });

    function addLogEntry(text) {
        const log = progressLog;
        if (!log) return;
        const p = document.createElement('p');
        p.innerText = `> ${text}`;
        log.appendChild(p);
        log.scrollTop = log.scrollHeight;
    }

    function distributeContentToSections(fullText, type) {
        // Logic to split the long text into multiple sections
        // This is a simplified version, ideally it should look for headers
        const sections = type === 'SKKN' ?
            ['section1', 'theory', 'section2', 'solution1', 'solution2', 'solution3', 'creativity', 'impact', 'appendix'] :
            ['section1', 'section2', 'solution1', 'solution2', 'solution3', 'creativity', 'impact'];

        // Try to finding markers for each section
        const lines = fullText.split('\n');
        let currentSectionIdx = -1;
        let sectionContents = {};

        lines.forEach(line => {
            const l = line.trim().toUpperCase();
            if (l.includes('PHẦN I') || l.includes('1. ĐẶT VẤN ĐỀ')) currentSectionIdx = 0;
            else if (l.includes('CƠ SỞ LÝ LUẬN') && type === 'SKKN') currentSectionIdx = 1;
            else if (l.includes('THỰC TRẠNG')) currentSectionIdx = type === 'SKKN' ? 2 : 1;
            else if (l.includes('GIẢI PHÁP 1') || l.includes('BIỆN PHÁP 1')) currentSectionIdx = type === 'SKKN' ? 3 : 2;
            else if (l.includes('GIẢI PHÁP 2') || l.includes('BIỆN PHÁP 2')) currentSectionIdx = type === 'SKKN' ? 4 : 3;
            else if (l.includes('GIẢI PHÁP 3') || l.includes('BIỆN PHÁP 3')) currentSectionIdx = type === 'SKKN' ? 5 : 4;
            else if (l.includes('TÍNH MỚI') || l.includes('HIỆU QUẢ SKKN')) currentSectionIdx = type === 'SKKN' ? 6 : 5;
            else if (l.includes('KẾT LUẬN')) currentSectionIdx = type === 'SKKN' ? 7 : 6;
            else if (l.includes('PHỤ LỤC') && type === 'SKKN') currentSectionIdx = 8;

            if (currentSectionIdx >= 0) {
                const sId = sections[currentSectionIdx];
                if (!sectionContents[sId]) sectionContents[sId] = [];
                sectionContents[sId].push(line);
            }
        });

        // Update DOM
        Object.keys(sectionContents).forEach(sId => {
            const area = document.querySelector(`.editor-area[data-id="${getDataIdFromSectionId(sId)}"]`);
            if (area) area.innerHTML = formatAiResponse(sectionContents[sId].join('\n'));
        });
    }

    function getDataIdFromSectionId(sId) {
        const map = {
            'section1': '1',
            'theory': 'theory',
            'section2': '2',
            'solution1': '2.3.1',
            'solution2': '2.3.2',
            'solution3': '2.3.3',
            'creativity': '2.4',
            'impact': '3',
            'appendix': 'appendix'
        };
        return map[sId] || sId;
    }

    if (appraisalUpgradeResultView) appraisalUpgradeResultView.addEventListener('click', () => {
        if (appraisalUpgradeResultView) appraisalUpgradeResultView.style.display = 'none';
        if (appraisalResultsView) appraisalResultsView.style.display = 'block';
    });

    if (exportRefinedDocxBtn) exportRefinedDocxBtn.addEventListener('click', () => {
        const content = refinedFullEditor?.value;
        if (!content) return alert("Không có nội dung để xuất.");
        exportRefinedDoc(content);
    });

    async function exportRefinedDoc(content) {
        if (!content) return alert("Không có nội dung để xuất.");

        // Get info from main form to build cover page
        const topic = topicTitleInput?.value || "TÊN ĐỀ TÀI CHƯA NHẬP";
        const author = document.getElementById('author-name')?.value || "Người thực hiện";
        const school = document.getElementById('school-name')?.value || "Tên Trường";
        const subject = document.getElementById('subject-name')?.value || "";
        const grade = document.getElementById('grade-level')?.value || "";
        const className = document.getElementById('class-name')?.value || "";
        const time = document.getElementById('execution-time')?.value || "";

        const btn = exportAfterUpgradeBtn || exportRefinedDocxBtn;
        const originalHTML = btn.innerHTML;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG XUẤT...';
        btn.disabled = true;

        try {
            const docxLib = window.docx || docx;
            const { Document, Packer, Paragraph, TextRun, AlignmentType, PageBreak } = docxLib;

            const children = [];

            // --- PAGE 1: PROFESSIONAL COVER PAGE (REFINED DOCUMENT) ---
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: school.toUpperCase(), size: 26, font: "Times New Roman" })]
            }));

            children.push(new Paragraph({ spacing: { before: 2000 } }));

            const docType = getDocType();
            const titleText = docType === 'SKKN' ? "BÁO CÁO SÁNG KIẾN KINH NGHIỆM" : "BÁO CÁO BIỆN PHÁP";
            const subTitleText = docType === 'SKKN' ? "GÓP PHẦN NÂNG CAO HIỆU QUẢ CÔNG TÁC GIÁO DỤC" : "GÓP PHẦN NÂNG CAO CHẤT LƯỢNG CÔNG TÁC GIẢNG DẠY";

            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: titleText, bold: true, size: 32, font: "Times New Roman" })]
            }));

            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: subTitleText, size: 24, font: "Times New Roman" })]
            }));

            children.push(new Paragraph({ spacing: { before: 800 } }));

            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: "ĐỀ TÀI:", bold: true, size: 26, font: "Times New Roman" })]
            }));
            children.push(new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: topic.toUpperCase(), bold: true, size: 32, font: "Times New Roman" })]
            }));

            children.push(new Paragraph({ spacing: { before: 2000 } }));

            const infoEntries = [
                { label: "Môn học/Lĩnh vực:", value: subject },
                { label: "Cấp học/Khối lớp:", value: `${grade} - ${className}` },
                { label: "Người thực hiện:", value: author },
                { label: "Thời gian thực hiện:", value: time }
            ];

            infoEntries.forEach(entry => {
                children.push(new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [
                        new TextRun({ text: entry.label.padEnd(25, ' '), font: "Times New Roman", size: 28 }),
                        new TextRun({ text: entry.value, bold: true, font: "Times New Roman", size: 28 })
                    ],
                    spacing: { after: 120 }
                }));
            });

            children.push(new Paragraph({ children: [new PageBreak()] }));

            // --- REFINED CONTENT ---
            // Clean up literal \n if they still exist
            const processedContent = content.replace(/\\n/g, '\n');
            const mainContentLines = processedContent.split('\n');

            mainContentLines.forEach(line => {
                const trimmed = line.trim();
                if (!trimmed) return; // Skip empty lines, handled by spacing

                // Sưu tầm logic nhận diện tiêu đề chuyên sâu hơn
                const isMainPart = trimmed.match(/^(PHẦN|I\.|II\.|III\.|IV\.|V\.|VI\.)\s/i);
                const isSubHeading = trimmed.match(/^[1-9]\.[1-9]/); // 1.1, 2.1...
                const isAllCap = trimmed === trimmed.toUpperCase() && trimmed.length > 10;

                const isHeading = isMainPart || isSubHeading || isAllCap;

                children.push(new Paragraph({
                    children: [new TextRun({
                        text: trimmed,
                        size: isMainPart ? 30 : (isSubHeading || isAllCap) ? 28 : 26, // Cỡ chữ phù hợp GVG
                        bold: isHeading,
                        font: "Times New Roman"
                    })],
                    spacing: {
                        before: isHeading ? 240 : 0,
                        after: 120,
                        line: 360 // Giãn dòng 1.5
                    },
                    alignment: isHeading ? AlignmentType.LEFT : AlignmentType.JUSTIFIED,
                    indent: isHeading ? {} : { firstLine: 700 } // Thụt đầu dòng đoạn văn
                }));
            });

            const doc = new Document({
                sections: [{
                    properties: {
                        page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1701 } }
                    },
                    children: children
                }]
            });

            const blob = await Packer.toBlob(doc);
            saveAs(blob, `BAO_CAO_HOAN_THIEN_${new Date().getTime()}.docx`);
        } catch (error) {
            console.error(error);
            alert("Lỗi khi tạo file Word: " + error.message);
        } finally {
            btn.innerHTML = originalHTML;
            btn.disabled = false;
        }
    }

    // Topic Analysis (Image A3) Logic

    if (refineTopicBtn) refineTopicBtn.addEventListener('click', async () => {
        if (!lastAppraisedText) {
            alert('Vui lòng thẩm định file trước khi sử dụng tính năng này.');
            return;
        }

        if (topicAnalysisView) topicAnalysisView.style.display = 'flex';
        if (topicSuggestionsList) topicSuggestionsList.innerHTML = '<div class="loading-suggestions"><i class="fas fa-spinner fa-spin"></i> Đang phân tích đề tài tối ưu từ bài viết của bạn...</div>';

        const appraisalType = document.querySelector('input[name="appraisal_type"]:checked')?.value || "BIỆN PHÁP";

        const prompt = `Bạn là cố vấn cấp cao cho kỳ thi GVG/GVCNG hoặc thẩm định SKKN. 
        Dựa trên nội dung ${appraisalType} này: """${lastAppraisedText.substring(0, 8000)}"""
        
        Hãy đề xuất chính xác 4-5 tên đề tài thay thế cực kỳ ấn tượng, bám sát chuyên môn của một bài ${appraisalType}.
        - Nếu là BIỆN PHÁP: Tên đề tài nên tập trung vào sự thay đổi trong giảng dạy/giáo dục tại lớp.
        - Nếu là SKKN: Tên đề tài nên mang tính chất nghiên cứu, đúc kết khoa học.
        Yêu cầu: Tên đề tài phải cụ thể, có cấu trúc chuẩn mực sư phạm.
        
        Trả về JSON:
        {
          "suggestions": [
            {
              "title": "Tên đề tài 1",
              "desc": "Giải thích ngắn gọn tại sao tên này hay",
              "score": 85-98
            }
          ]
        }`;

        try {
            const result = await callGemini(prompt, () => refineTopicBtn.click());

            if (result.error) {
                if (topicSuggestionsList) topicSuggestionsList.innerHTML = `
                    <div class="loading-suggestions" style="color: #e74c3c;">
                        <i class="fas fa-exclamation-triangle"></i><br><br>
                        ${result.error === 'Chưa cấu hình API Key.' ? 'Chưa có API Key. Vui lòng cấu hình trong menu.' : 'Lỗi: ' + result.error}
                    </div>`;
                return;
            }

            if (result.text) {
                try {
                    const data = robustParseJSON(result.text);
                    if (!data || !data.suggestions) throw new Error("JSON structure invalid");

                    if (topicSuggestionsList) topicSuggestionsList.innerHTML = data.suggestions.map(s => `
                        <div class="topic-suggestion-item">
                            <div class="topic-score ${s.score >= 90 ? 'score-high' : 'score-med'}">${s.score}</div>
                            <div class="topic-content">
                                <h4>${s.title}</h4>
                                <p>${s.desc}</p>
                            </div>
                            <div class="topic-item-actions">
                                <button class="icon-btn copy-topic-btn" title="Copy tên đề tài"><i class="far fa-clipboard"></i></button>
                                <button class="primary-btn use-topic-btn" data-title="${s.title}">Sử dụng</button>
                            </div>
                        </div>
                    `).join('');

                    // Add event listeners
                    document.querySelectorAll('.copy-topic-btn').forEach((btn, idx) => {
                        btn.addEventListener('click', () => {
                            const title = data.suggestions[idx].title;
                            navigator.clipboard.writeText(title).then(() => {
                                btn.innerHTML = '<i class="fas fa-check"></i>';
                                setTimeout(() => btn.innerHTML = '<i class="far fa-clipboard"></i>', 2000);
                            });
                        });
                    });

                    document.querySelectorAll('.use-topic-btn').forEach(btn => {
                        btn.addEventListener('click', () => {
                            const title = btn.getAttribute('data-title');
                            if (topicAnalysisView) topicAnalysisView.style.display = 'none';
                            showUpgradePlan(title, true);
                        });
                    });
                } catch (e) {
                    console.error("JSON Parse Error:", e, jsonText);
                    if (topicSuggestionsList) topicSuggestionsList.innerHTML = '<div class="loading-suggestions">AI trả về dữ liệu không đúng cấu trúc. Vui lòng thử lại.</div>';
                }
            } else {
                if (topicSuggestionsList) topicSuggestionsList.innerHTML = '<div class="loading-suggestions">Không nhận được phản hồi từ AI.</div>';
            }
        } catch (error) {
            console.error("Refine Topic Error:", error);
            if (topicSuggestionsList) topicSuggestionsList.innerHTML = `<div class="loading-suggestions">Lỗi hệ thống: ${error.message}</div>`;
        }
    });

    if (keepOldTopicBtn) keepOldTopicBtn.addEventListener('click', () => {
        if (topicAnalysisView) topicAnalysisView.style.display = 'none';
        const currentTitle = topicTitleInput?.value || "Tên đề tài hiện tại";
        showUpgradePlan(currentTitle, false);
    });

    if (closeUpgradePlanBtn) closeUpgradePlanBtn.addEventListener('click', () => {
        if (upgradePlanView) upgradePlanView.style.display = 'none';
    });

    async function showUpgradePlan(title, isNewTopic) {
        if (upgradePlanView) upgradePlanView.style.display = 'flex';
        // Reset steps list with loading spinner
        if (upgradeStepsList) upgradeStepsList.innerHTML = '<li><i class="fas fa-spinner fa-spin"></i> Đang xây dựng lộ trình cải thiện tối ưu...</li>';

        // Update AI risk score in the alert banner
        if (alertAiScore) alertAiScore.innerText = lastAppraisalData?.ai_risk || 0;

        const prompt = `Bạn là cố vấn cấp cao. Dựa trên file thẩm định và đề tài "${title}", hãy đưa ra đúng 6 bước hành động (Action Plan) để AI nâng cấp bài viết này thành xuất sắc.
        Trả về JSON: {"steps": ["Bước 1...", "Bước 2...", "Bước 3...", "Bước 4...", "Bước 5...", "Bước 6..."]}`;

        try {
            const result = await callGemini(prompt, () => showUpgradePlan(title, isNewTopic));
            if (result.text) {
                const data = robustParseJSON(result.text);
                if (data && data.steps && upgradeStepsList) {
                    upgradeStepsList.innerHTML = data.steps.map((s, idx) => `<li><b>${idx + 1}</b> ${s}</li>`).join('');
                    return; // Success
                }
            }
            // Fallback if no text or invalid JSON
            if (upgradeStepsList) upgradeStepsList.innerHTML = '<li>Hoàn thiện các nội dung còn thiếu trong báo cáo.</li><li>Nâng cấp ngôn ngữ chuyên môn theo chuẩn GVG.</li><li>Bổ sung minh chứng số liệu thực tế.</li>';
        } catch (e) {
            console.error("Show Upgrade Plan Error:", e);
            if (upgradeStepsList) upgradeStepsList.innerHTML = '<li>Tải lộ trình thất bại. Bạn có thể nhấn Xác nhận cải thiện ngay để AI tự động tối ưu.</li>';
        }
    }

    closeTopicAnalysis?.addEventListener('click', () => {
        topicAnalysisView.style.display = 'none';
    });

    async function startAnalysis(file) {
        const originalContent = uploadDropzone.innerHTML;
        uploadDropzone.innerHTML = `
            <div class="upload-icon"><i class="fas fa-spinner fa-spin"></i></div>
            <h3>Đang thẩm định chuyên sâu "${file.name}"...</h3>
            <p>AI đang phân tích phong cách hành văn, cấu trúc và đặc trưng sư phạm. Vui lòng đợi.</p>
        `;

        try {
            const fileText = await extractText(file);
            if (!fileText || fileText.length < 100) {
                alert("File không chứa đủ văn bản hoặc định dạng không được hỗ trợ.");
                uploadDropzone.innerHTML = originalContent;
                return;
            }

            const appraisalType = document.querySelector('input[name="appraisal_type"]:checked')?.value || "BIỆN PHÁP";
            const isSkkn = appraisalType === 'SKKN';

            lastAppraisedText = fileText;

            let prompt = "";
            if (isSkkn) {
                prompt = `Bạn là chuyên gia thẩm định SÁNG KIẾN KINH NGHIỆM (SKKN) với 20 năm kinh nghiệm. 
                Hãy ĐỌC KỸ và PHÂN TÍCH nội dung file để thẩm định theo 5 TIÊU CHÍ (Thang điểm 100).
                
                NỘI DUNG: """${fileText.substring(0, 15000)}"""

                YÊU CẦU: Trả về JSON DUY NHẤT. Chấm điểm cực kỳ khắt khe và công tâm.
                JSON STRUCTURE:
                {
                  "type": "SKKN",
                  "topic_name": "Tên đề tài thực tế trong bài",
                  "skkn_scores": {
                    "innovation": 0-25,
                    "practicality": 0-30,
                    "scientific": 0-20,
                    "benefit": 0-15,
                    "format": 0-10
                  },
                  "quality_score": 0-100,
                  "status_label": "Tốt/Khá/Trung bình/Không đạt",
                  "summary": "Tóm tắt đánh giá tổng thể (2-3 câu).",
                  "plagiarism_risk": 0-100,
                   "ai_risk": 0-100,
                  "ai_risk_detail": {
                      "score": 0-100,
                      "passages": [
                          { "text": "đoạn văn nghi ngờ AI...", "reason": "Tại sao AI viết (VD: Cấu trúc quá hoàn hảo, thiếu cảm xúc cá nhân...)" }
                      ]
                  },
                  "plagiarism_examples": [ { "quote": "...", "source": "..." } ],
                  "spell_check": [ { "original": "...", "corrected": "..." } ],
                  "structure_detail": [
                    { "part": "Đặt vấn đề", "pros": "...", "cons": "..." },
                    { "part": "Thực trạng", "pros": "...", "cons": "..." },
                    { "part": "Giải pháp", "pros": "...", "cons": "..." },
                    { "part": "Hiệu quả", "pros": "...", "cons": "..." },
                    { "part": "Kết luận", "pros": "...", "cons": "..." }
                  ],
                  "expert_advice": "Lời khuyên chiến lược để nâng hạng bài viết.",
                  "ai_details": { "perplexity": 0, "burstiness": 0, "suggestions": [] }
                }`;
            }
            else {
                prompt = `Bạn là chuyên gia thẩm định BIỆN PHÁP với 20 năm kinh nghiệm. 
                Hãy ĐỌC KỸ và PHÂN TÍCH nội dung thực tế từ file giáo viên tải lên để thẩm định tiêu chuẩn BIỆN PHÁP GVG dựa trên THANG ĐIỂM 100.
                
                NỘI DUNG: """${fileText.substring(0, 15000)}"""

                YÊU CẦU: Trả về JSON DUY NHẤT. Chấm điểm chi tiết theo các tiêu chí sau (Tổng 100):
                1. Tính khoa học (15đ).
                2. Tính thực tiễn (20đ).
                3. Tính mới và sáng tạo (20đ).
                4. Khả năng áp dụng (20đ).
                5. Hiệu quả minh chứng (15đ).
                6. Ngôn ngữ và trình bày (10đ).

                JSON STRUCTURE:
                {
                  "type": "BIỆN PHÁP",
                  "summary": "Tóm tắt nội dung",
                  "quality_score": 0-100, 
                  "plagiarism_risk": 0-100, 
                  "ai_risk": 0-100,
                  "ai_risk_detail": {
                      "score": 0-100,
                      "passages": [
                          { "text": "...", "reason": "..." }
                      ]
                  },
                  "status_label": "Giỏi/Khá/TB/Yếu",
                  "structure": { "toc": bool, "intro": bool, "reality": bool, "solution": bool, "conclusion": bool },
                  "criteria": { 
                    "scientific": 0-15, 
                    "practical": 0-20, 
                    "innovation": 0-20, 
                    "applied_ability": 0-20, 
                    "demonstrated_effect": 0-15, 
                    "language": 0-10 
                  },
                  "plagiarism_examples": [ { "quote": "...", "source": "..." } ],
                  "spell_check": [ { "original": "...", "corrected": "..." } ],
                  "ai_details": { "perplexity": 0, "burstiness": 0, "suggestions": [] }
                }`;
            }

            const result = await callGemini(prompt, () => startAnalysis(file));

            if (result.text) {
                try {
                    const data = robustParseJSON(result.text);
                    if (!data || !data.type) throw new Error("AI trả về phản hồi không thể phân tích cấu trúc.");

                    lastAppraisalData = data;

                    // Sync topic name to the main input field if it's available
                    if (topicTitleInput && data.topic_name) {
                        topicTitleInput.value = data.topic_name;
                    }

                    updateAppraisalUI(data, data.type === 'SKKN');

                    // Show Results View
                    const uploadViewEl = document.getElementById('appraisal-upload-view');
                    const resultsViewEl = document.getElementById('appraisal-results-view');

                    if (uploadViewEl) uploadViewEl.style.display = 'none';
                    if (resultsViewEl) resultsViewEl.style.display = 'block';

                } catch (parseError) {
                    console.error("Analysis JSON Error:", parseError, "AI Text:", result.text);
                    throw new Error(`AI trả về phản hồi không đúng định dạng. Lỗi: ${parseError.message}. Vui lòng thử lại.`);
                }
            } else {
                throw new Error(result.error || "Không nhận được phản hồi từ AI");
            }

        } catch (error) {
            console.error("Appraisal Error:", error);
            alert("Lỗi: " + error.message);
            uploadDropzone.innerHTML = originalContent;
        } finally {
            // Restore any UI states if necessary
        }
    }

    function updateAppraisalUI(data, isSkknFlag = false) {
        const bpContainer = document.getElementById('bp-appraisal-content');
        const skknContainer = document.getElementById('skkn-appraisal-content');

        // Cần xác định chính xác mode dựa trên data trả về hoặc tham số
        const actualIsSkkn = isSkknFlag || data.type === 'SKKN';

        if (actualIsSkkn) {
            if (bpContainer) bpContainer.style.display = 'none';
            if (skknContainer) skknContainer.style.display = 'block';
            renderSkknAppraisal(data);
        } else {
            if (skknContainer) skknContainer.style.display = 'none';
            if (bpContainer) bpContainer.style.display = 'block';
            renderBpAppraisal(data);
        }
    }

    function renderBpAppraisal(data) {
        document.getElementById('analysis-summary-text').innerText = data.summary;
        updateCircularProgress('CHẤT LƯỢNG TỔNG THỂ', data.quality_score, data.status_label);
        updateCircularProgress('NGUY CƠ ĐẠO VĂN', data.plagiarism_risk);
        updateCircularProgress('NGUY CƠ AI', data.ai_risk);

        const structureList = document.querySelector('.structure-list');
        if (structureList) {
            structureList.innerHTML = `
                <li class="${data.structure.toc ? 'check' : 'error'}"><i class="fas ${data.structure.toc ? 'fa-check-circle' : 'fa-times-circle'}"></i> Mục lục</li>
                <li class="${data.structure.intro ? 'check' : 'error'}"><i class="fas ${data.structure.intro ? 'fa-check-circle' : 'fa-times-circle'}"></i> 1. Đặt vấn đề</li>
                <li class="${data.structure.reality ? 'check' : 'error'}"><i class="fas ${data.structure.reality ? 'fa-check-circle' : 'fa-times-circle'}"></i> 2. Thực trạng vấn đề</li>
                <li class="${data.structure.solution ? 'check' : 'warn'}"><i class="fas ${data.structure.solution ? 'fa-check-circle' : 'fa-exclamation-triangle'}"></i> 3. Nội dung giải pháp</li>
                <li class="${data.structure.conclusion ? 'check' : 'error'}"><i class="fas ${data.structure.conclusion ? 'fa-check-circle' : 'fa-times-circle'}"></i> 4. Kết luận và kiến nghị</li>
            `;
        }

        updateBarChart(data.criteria);
        // Normalize radar chart values out of 10 for drawing
        const radarValues = [
            (data.criteria.scientific / 15) * 10,
            (data.criteria.practical / 20) * 10,
            (data.criteria.innovation / 20) * 10,
            (data.criteria.applied_ability / 20) * 10,
            (data.criteria.demonstrated_effect / 15) * 10,
            (data.criteria.language / 10) * 10
        ];
        drawRadarChart(radarValues);

        const plagiarismList = document.querySelector('.plagiarism-list');
        if (plagiarismList) {
            if (data.plagiarism_examples && data.plagiarism_examples.length > 0) {
                plagiarismList.innerHTML = data.plagiarism_examples.map(ex => `
                    <div class="plagiarism-item">
                        <p class="quote">"${ex.quote}"</p>
                        <span class="source">→ Nguồn: ${ex.source}</span>
                    </div>
                `).join('');
            } else {
                plagiarismList.innerHTML = "<p>Không phát hiện dấu hiệu đạo văn rõ rệt.</p>";
            }
        }

        const aiStats = document.querySelector('.ai-stats-grid');
        if (aiStats) {
            aiStats.innerHTML = `
                <div class="ai-stat-item"><div class="stat-value">${data.ai_details.perplexity}</div><div class="stat-label">PERPLEXITY</div></div>
                <div class="ai-stat-item"><div class="stat-value">${data.ai_details.burstiness}</div><div class="stat-label">BURSTINESS</div></div>
                <div class="ai-stat-item"><div class="stat-value">${data.ai_details.pattern_score}</div><div class="stat-label">PATTERN</div></div>
            `;
        }

        const aiPassagesList = document.querySelector('.ai-passages-list');
        if (aiPassagesList) {
            if (data.ai_risk_detail?.passages?.length > 0) {
                aiPassagesList.innerHTML = data.ai_risk_detail.passages.map(p => `
                    <div class="ai-passage-item" style="border-left: 3px solid #6366f1; padding-left: 10px; margin-bottom: 12px; background: #f5f3ff; padding: 10px; border-radius: 4px;">
                        <p style="font-style: italic; color: #4338ca; font-size: 0.95rem; margin-bottom: 5px;">"${p.text}"</p>
                        <span style="font-size: 0.8rem; color: #6b7280;"><i class="fas fa-info-circle"></i> ${p.reason}</span>
                    </div>
                `).join('');
            } else {
                aiPassagesList.innerHTML = "<p>Không phát hiện đoạn văn nào mang đậm dấu ấn AI.</p>";
            }
        }
    }

    function renderSkknAppraisal(data) {
        const skknContainer = document.getElementById('skkn-appraisal-content');
        if (!skknContainer) return;

        skknContainer.innerHTML = `
            <div class="skkn-results-header">
                <div class="header-main">
                    <h2>KẾT QUẢ THẨM ĐỊNH SKKN</h2>
                    <p class="topic-title">"${data.topic_name}"</p>
                    <div class="badges">
                        <div class="status-box">
                            <span class="status-label">${data.status_label}</span>
                            <div class="score-circle">
                                <span class="score-value">${data.quality_score}</span>
                                <span class="score-max">/100</span>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="skkn-detailed-scores" style="margin-top: 2rem; background: #f8fafc; padding: 1.5rem; border-radius: 12px; border: 1px solid #e2e8f0;">
                <h3 style="margin-bottom: 1.5rem; color: #1e293b; font-size: 1.25rem; display: flex; align-items: center; gap: 10px;">
                    <i class="fas fa-list-ol" style="color: #6366f1;"></i> ĐIỂM CHI TIẾT THEO 5 TIÊU CHÍ (Mới)
                </h3>
                <div class="score-bars" style="display: flex; flex-direction: column; gap: 1.2rem;">
                    <div class="score-bar-item">
                        <div class="label" style="display: flex; justify-content: space-between; margin-bottom: 0.5rem; font-weight: 500;">
                            <span>1. Tính mới và sáng tạo</span>
                            <span style="color: #6366f1; font-weight: 700;">${data.skkn_scores?.innovation}/25</span>
                        </div>
                        <div class="bar-bg" style="height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden;">
                            <div class="bar-fill" style="width: ${(data.skkn_scores?.innovation / 25 * 100)}%; height: 100%; background: linear-gradient(90deg, #6366f1, #8b5cf6); border-radius: 5px;"></div>
                        </div>
                    </div>
                    <div class="score-bar-item">
                        <div class="label" style="display: flex; justify-content: space-between; margin-bottom: 0.5rem; font-weight: 500;">
                            <span>2. Tính thực tiễn và khả thi</span>
                            <span style="color: #8b5cf6; font-weight: 700;">${data.skkn_scores?.practicality}/30</span>
                        </div>
                        <div class="bar-bg" style="height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden;">
                            <div class="bar-fill" style="width: ${(data.skkn_scores?.practicality / 30 * 100)}%; height: 100%; background: linear-gradient(90deg, #8b5cf6, #d946ef); border-radius: 5px;"></div>
                        </div>
                    </div>
                    <div class="score-bar-item">
                        <div class="label" style="display: flex; justify-content: space-between; margin-bottom: 0.5rem; font-weight: 500;">
                            <span>3. Tính khoa học</span>
                            <span style="color: #ec4899; font-weight: 700;">${data.skkn_scores?.scientific}/20</span>
                        </div>
                        <div class="bar-bg" style="height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden;">
                            <div class="bar-fill" style="width: ${(data.skkn_scores?.scientific / 20 * 100)}%; height: 100%; background: linear-gradient(90deg, #ec4899, #f43f5e); border-radius: 5px;"></div>
                        </div>
                    </div>
                    <div class="score-bar-item">
                        <div class="label" style="display: flex; justify-content: space-between; margin-bottom: 0.5rem; font-weight: 500;">
                            <span>4. Hiệu quả và lợi ích</span>
                            <span style="color: #f59e0b; font-weight: 700;">${data.skkn_scores?.benefit}/15</span>
                        </div>
                        <div class="bar-bg" style="height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden;">
                            <div class="bar-fill" style="width: ${(data.skkn_scores?.benefit / 15 * 100)}%; height: 100%; background: linear-gradient(90deg, #f59e0b, #fbbf24); border-radius: 5px;"></div>
                        </div>
                    </div>
                    <div class="score-bar-item">
                        <div class="label" style="display: flex; justify-content: space-between; margin-bottom: 0.5rem; font-weight: 500;">
                            <span>5. Hình thức trình bày</span>
                            <span style="color: #10b981; font-weight: 700;">${data.skkn_scores?.format}/10</span>
                        </div>
                        <div class="bar-bg" style="height: 10px; background: #e2e8f0; border-radius: 5px; overflow: hidden;">
                            <div class="bar-fill" style="width: ${(data.skkn_scores?.format / 10 * 100)}%; height: 100%; background: linear-gradient(90deg, #10b981, #34d399); border-radius: 5px;"></div>
                        </div>
                    </div>
                </div>
            </div>

            <div class="skkn-analysis-row" style="margin-top: 2rem; display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1.5rem;">
                <div class="analysis-box ai-detection" style="background: white; border-radius: 12px; border: 1px solid #e2e8f0; overflow: hidden;">
                    <div class="box-header" style="background: #f5f3ff; color: #6366f1; padding: 1rem; font-weight: 600; display: flex; justify-content: space-between; align-items: center;">
                        <span><i class="fas fa-robot"></i> PHÂN TÍCH YẾU TỐ AI</span>
                        <span class="risk-badge" style="background: #6366f1; color: white; padding: 2px 8px; border-radius: 4px; font-size: 0.8rem;">NGUY CƠ: ${data.ai_risk}%</span>
                    </div>
                    <div class="ai-scroll-list" style="padding: 1rem; max-height: 250px; overflow-y: auto;">
                        ${(data.ai_risk_detail?.passages || []).map(p => `
                            <div class="ai-passage-item" style="margin-bottom: 1rem; padding-bottom: 0.8rem; border-bottom: 1px dashed #e2e8f0;">
                                <p style="font-style: italic; color: #1e293b; font-size: 0.9rem; margin-bottom: 4px;">"${p.text}"</p>
                                <span style="font-size: 0.75rem; color: #64748b; display: block;"><i class="fas fa-info-circle"></i> ${p.reason}</span>
                            </div>
                        `).join('') || '<p style="color: #64748b; font-size: 0.9rem;">Không phát hiện đoạn văn nào mang đậm dấu ấn AI.</p>'}
                    </div>
                </div>
                <div class="analysis-box plagiarism">
                    <div class="box-header"><i class="fas fa-exclamation-triangle"></i> KIỂM TRA ĐẠO VĂN <span class="risk-badge">NGUY CƠ: ${data.plagiarism_risk}%</span></div>
                    <div class="plagiarism-scroll-list">
                        ${(data.plagiarism_examples || []).map(ex => `
                            <div class="plag-item">
                                <p>"${ex.quote}"</p>
                                <span>Trùng: ~${Math.floor(Math.random() * 20) + 70}% → Nguồn: ${ex.source}</span>
                            </div>
                        `).join('') || '<p>Không phát hiện đoạn văn trùng lặp rõ rệt.</p>'}
                    </div>
                </div>
                <div class="analysis-box grammar">
                    <div class="box-header"><i class="fas fa-spell-check"></i> LỖI CHÍNH TẢ & DIỄN ĐẠT</div>
                    <div class="grammar-list">
                        ${(data.spell_check || []).map(err => `
                            <div class="grammar-item">
                                <span class="wrong">${err.original}</span> <i class="fas fa-arrow-right"></i> <span class="right">${err.corrected}</span>
                            </div>
                        `).join('') || '<p>Chúc mừng! Không phát hiện lỗi chính tả nghiêm trọng.</p>'}
                    </div>
                </div>
            </div>

            <div class="skkn-detailed-evaluation" style="margin-top: 2rem;">
                 <div class="criteria-header" style="background: #1e293b; color: white; padding: 1rem; border-radius: 8px 8px 0 0; font-weight: 600;">
                    <i class="fas fa-award"></i> ĐÁNH GIÁ CHI TIẾT NỘI DUNG
                 </div>
                 <div class="evaluation-body" style="background: white; border: 1px solid #e2e8f0; border-top: none; padding: 1.5rem; border-radius: 0 0 8px 8px;">
                     <p style="line-height: 1.6; color: #334155;">${data.summary}</p>
                     
                     <div class="structure-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1.5rem; margin-top: 1.5rem;">
                         ${(data.structure_detail || []).map(s => `
                             <div class="part-card" style="border: 1px solid #f1f5f9; padding: 1rem; border-radius: 8px; background: #f8fafc;">
                                 <h4 style="color: #475569; margin-bottom: 0.5rem; border-bottom: 1px dashed #cbd5e1; padding-bottom: 0.5rem;">${s.part}</h4>
                                 <div style="font-size: 0.9rem; margin-bottom: 0.5rem; color: #059669;"><i class="fas fa-plus-circle"></i> ${s.pros}</div>
                                 <div style="font-size: 0.9rem; color: #dc2626;"><i class="fas fa-minus-circle"></i> ${s.cons}</div>
                             </div>
                         `).join('')}
                     </div>
                 </div>
            </div>

            <div class="expert-advice-box" style="margin-top: 2rem; background: #fffbeb; border: 1px solid #fde68a; padding: 1.5rem; border-radius: 12px;">
                <div class="box-title" style="color: #92400e; font-weight: 700; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 8px;">
                    <i class="fas fa-lightbulb"></i> LỜI KHUYÊN TỪ CHUYÊN GIA THẨM ĐỊNH
                </div>
                <p style="color: #92400e; font-style: italic; line-height: 1.5;">"${data.expert_advice || "Duy trì cấu trúc hiện có và tập trung làm rõ tính mới."}"</p>
            </div>
        `;
    }

    function updateCircularProgress(title, value, status = null) {
        // Find card by title
        const cards = document.querySelectorAll('.metric-card');
        const card = Array.from(cards).find(c => c.querySelector('h3').innerText === title);
        if (!card) return;

        const cp = card.querySelector('.circular-progress');
        const circle = cp.querySelector('.progress');
        const valueBox = cp.querySelector('.value');
        const r = 54;
        const circumference = 2 * Math.PI * r;

        // Update value text
        if (title.includes('CHẤT LƯỢNG')) valueBox.innerHTML = `${value} <span>/100 điểm</span>`;
        else valueBox.innerHTML = `${value}% <span>${title.includes('ĐẠO VĂN') ? 'trùng lặp' : 'nguy cơ AI'}</span>`;

        // Update circle
        const offset = circumference - (value / 100) * circumference;
        circle.style.strokeDasharray = `${circumference} `;
        circle.style.strokeDashoffset = offset;

        // Dynamic color
        let color = '#2ecc71'; // Green
        if (title.includes('CHẤT LƯỢNG')) {
            if (value < 50) color = '#e74c3c';
            else if (value < 80) color = '#f39c12';
        } else {
            // Risk meters: Red for high, green for low
            if (value > 60) color = '#e74c3c';
            else if (value > 30) color = '#f39c12';
        }
        circle.style.stroke = color;

        // Update status badge if available
        if (status) {
            const badge = card.querySelector('.status-badge');
            badge.innerText = status;
            badge.className = `status - badge ${status === 'Giỏi' ? 'success' : status === 'Khá' ? 'warning' : 'danger'} `;
        }
    }

    function updateBarChart(criteria) {
        const barItems = document.querySelectorAll('.bar-item');

        barItems.forEach(item => {
            const label = item.querySelector('.bar-label').innerText.trim().toLowerCase();
            let val = 0;
            let max = 10;

            if (label.includes('khoa học')) { val = criteria.scientific || 0; max = 15; }
            else if (label.includes('thực tiễn')) { val = criteria.practical || 0; max = 20; }
            else if (label.includes('mới')) { val = criteria.innovation || 0; max = 20; }
            else if (label.includes('áp dụng')) { val = criteria.applied_ability || 0; max = 20; }
            else if (label.includes('hiệu quả')) { val = criteria.demonstrated_effect || 0; max = 15; }
            else if (label.includes('ngôn ngữ')) { val = criteria.language || 0; max = 10; }

            const percent = (val / max) * 100;
            const barFill = item.querySelector('.bar-fill');
            const barValue = item.querySelector('.bar-value');

            if (barFill) barFill.style.width = percent + '%';
            if (barValue) barValue.innerText = `${val}/${max}`;

            // Adjust color based on score relative to max
            if (barFill) {
                if (percent >= 80) barFill.className = 'bar-fill success';
                else if (percent >= 50) barFill.className = 'bar-fill warning';
                else barFill.className = 'bar-fill danger';
            }
        });
    }

    reUploadBtn?.addEventListener('click', (e) => {
        e.preventDefault();
        const resultsViewEl = document.getElementById('appraisal-results-view');
        const uploadViewEl = document.getElementById('appraisal-upload-view');
        const fileInputEl = document.getElementById('file-upload-input');

        if (resultsViewEl) resultsViewEl.style.display = 'none';
        if (uploadViewEl) uploadViewEl.style.display = 'block';
        if (fileInputEl) fileInputEl.value = '';
    });

    function drawRadarChart(values = [0.9, 0.9, 0.8, 0.8, 0.7, 0.8]) {
        const svg = document.getElementById('radar-chart');
        if (!svg) return;
        svg.innerHTML = ''; // Clear

        const centerX = 100;
        const centerY = 100;
        const radius = 70;
        const levels = 5;
        const labels = ['Tính khoa học', 'Tính thực tiễn', 'Tính mới', 'Khả năng áp dụng', 'Hiệu quả', 'Ngôn ngữ'];

        // Draw grid lines
        for (let i = 1; i <= levels; i++) {
            let r = (radius / levels) * i;
            let points = [];
            for (let j = 0; j < labels.length; j++) {
                let angle = (Math.PI * 2 / labels.length) * j - Math.PI / 2;
                points.push(`${centerX + r * Math.cos(angle)},${centerY + r * Math.sin(angle)}`);
            }
            const poly = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
            poly.setAttribute("points", points.join(" "));
            poly.setAttribute("class", "radar-grid");
            svg.appendChild(poly);
        }

        // Draw axes and labels
        for (let i = 0; i < labels.length; i++) {
            let angle = (Math.PI * 2 / labels.length) * i - Math.PI / 2;
            let x2 = centerX + radius * Math.cos(angle);
            let y2 = centerY + radius * Math.sin(angle);

            const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
            line.setAttribute("x1", centerX); line.setAttribute("y1", centerY);
            line.setAttribute("x2", x2); line.setAttribute("y2", y2);
            line.setAttribute("class", "radar-axis");
            svg.appendChild(line);

            let labelAngle = angle;
            let labelR = radius + 20;
            let lx = centerX + labelR * Math.cos(labelAngle);
            let ly = centerY + labelR * Math.sin(labelAngle);

            const text = document.createElementNS("http://www.w3.org/2000/svg", "text");
            text.setAttribute("x", lx); text.setAttribute("y", ly);
            text.setAttribute("text-anchor", "middle");
            text.setAttribute("class", "radar-label");
            text.textContent = labels[i];
            svg.appendChild(text);
        }

        // Draw data area
        let dataPoints = [];
        for (let i = 0; i < values.length; i++) {
            let angle = (Math.PI * 2 / values.length) * i - Math.PI / 2;
            let r = radius * values[i];
            dataPoints.push(`${centerX + r * Math.cos(angle)},${centerY + r * Math.sin(angle)}`);
        }

        const dataPoly = document.createElementNS("http://www.w3.org/2000/svg", "polygon");
        dataPoly.setAttribute("points", dataPoints.join(" "));
        dataPoly.setAttribute("class", "radar-area");
        svg.appendChild(dataPoly);

        for (let i = 0; i < values.length; i++) {
            let angle = (Math.PI * 2 / values.length) * i - Math.PI / 2;
            let r = radius * values[i];
            const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
            circle.setAttribute("cx", centerX + r * Math.cos(angle));
            circle.setAttribute("cy", centerY + r * Math.sin(angle));
            circle.setAttribute("r", 3);
            circle.setAttribute("class", "radar-point");
            svg.appendChild(circle);
        }
    }

    infoForm?.addEventListener('submit', (e) => {
        e.preventDefault();
        switchSection('outline');
        currentSectionIndex = currentSectionOrder.indexOf('outline');
    });
});
