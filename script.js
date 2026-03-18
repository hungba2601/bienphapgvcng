document.addEventListener('DOMContentLoaded', () => {
    // LOGIN LOGIC
    const loginOverlay = document.getElementById('login-overlay');
    const appContainer = document.querySelector('.app-container');
    const loginBtn = document.getElementById('login-btn');
    const loginUsernameInput = document.getElementById('login-username');
    const loginPasswordInput = document.getElementById('login-password');
    const loginMessage = document.getElementById('login-message');

    // Link script.google.com của bạn (Phải cập nhật link này sau khi triển khai Apps Script)
    const SHEETS_API_URL = "https://script.google.com/macros/s/AKfycbzmZsB6Lnbc7kZjddMi6Gu5-Snn31LhmlcklFOtrl73ofXfjGrfoJUrTtGzb8t8ZF7c/exec";

    // Kiểm tra trạng thái đã đăng nhập chưa
    if (localStorage.getItem('isLoggedIn') === 'true') {
        const savedUser = localStorage.getItem('username') || 'Thành viên';
        const displayUser = document.getElementById('display-username');
        if (displayUser) displayUser.textContent = savedUser;

        loginOverlay.style.display = 'none';
        appContainer.style.display = 'flex';
    }

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

        try {
            const response = await fetch(`${SHEETS_API_URL}?action=login&user=${encodeURIComponent(user)}&pass=${encodeURIComponent(pass)}`);

            if (!response.ok) {
                throw new Error('Server trả về lỗi: ' + response.status);
            }

            const text = await response.text();
            let result;
            try {
                result = JSON.parse(text);
            } catch (e) {
                console.error("Phản hồi không phải JSON:", text);
                throw new Error('Dữ liệu từ Google Sheet không đúng định dạng. Hãy kiểm tra lại URL hoặc mã Script.');
            }

            if (result.status === 'success') {
                showLoginMessage('Đăng nhập thành công! Đang vào hệ thống...', 'success');
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
            showLoginMessage('Lỗi hệ thống: ' + error.message + '. (Hãy chắc chắn chọn "Anyone" khi Deploy bản mới nhất)', 'error');
            loginBtn.disabled = false;
            loginBtn.innerHTML = '<span>ĐĂNG NHẬP NGAY</span><i class="fas fa-sign-in-alt"></i>';
        }
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

    function cleanFileName(str) {
        if (!str) return "Bao_Cao_GVG";
        return str.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/g, 'd').replace(/Đ/g, 'D').replace(/[^a-zA-Z0-9]/g, '_').replace(/_+/g, '_').trim();
    }

    const sectionOrder = ['info', 'outline', 'section1', 'section2', 'solution1', 'solution2', 'solution3', 'creativity', 'impact'];
    let currentSectionIndex = 0;

    tabButtons.forEach(btn => {
        btn.addEventListener('click', () => {
            tabButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            if (btn.innerText.includes('THẨM ĐỊNH')) {
                sidebar.style.opacity = '0.5'; sidebar.style.pointerEvents = 'none';
                hideAllSections(); document.getElementById('appraisal-section').classList.add('active');
            } else {
                sidebar.style.opacity = '1'; sidebar.style.pointerEvents = 'auto';
                switchSection(sectionOrder[currentSectionIndex]);
            }
        });
    });

    navItems.forEach(item => {
        item.addEventListener('click', () => {
            const target = item.getAttribute('data-section');
            if (target) {
                switchSection(target);
                currentSectionIndex = sectionOrder.indexOf(target);
            }
        });
    });

    nextBtns.forEach(btn => btn.addEventListener('click', () => { if (currentSectionIndex < sectionOrder.length - 1) { currentSectionIndex++; switchSection(sectionOrder[currentSectionIndex]); } }));
    prevBtns.forEach(btn => btn.addEventListener('click', () => { if (currentSectionIndex > 0) { currentSectionIndex--; switchSection(sectionOrder[currentSectionIndex]); } }));

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
            const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${key}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ contents: [{ parts: [{ text: prompt + "\nVăn bản thuần, dùng ** cho in đậm tiêu đề." }] }] })
            });
            if (res.status === 429) {
                apiModal.classList.add('active');
                apiStatus.innerText = 'Hết hạn mức API. Vui lòng đổi Key khác.';
                if (onQuotaError) pendingAction = onQuotaError;
                return { error: 'Hết hạn mức (Quota).' };
            }
            const data = await res.json();
            let text = data.candidates?.[0]?.content?.parts?.[0]?.text;
            if (text) return { text: text.trim().replace(/^\s*\* /gm, '') };
            return { error: 'Lỗi AI' };
        } catch (e) { return { error: 'Lỗi kết nối' }; }
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
        const role = document.querySelector('input[name="teacher_role"]:checked')?.value;
        return role === 'GVCNG' ? "Giáo viên chủ nhiệm giỏi (GVCNG)" : "Giáo viên giỏi (GVG)";
    }

    function getDocType() {
        // Kiểm tra xem Tab "THẨM ĐỊNH" có đang active hay không
        const activeTab = document.querySelector('.tab-btn.active');
        if (activeTab && activeTab.innerText.includes('THẨM ĐỊNH')) {
            return document.querySelector('input[name="appraisal_type"]:checked')?.value || "BIỆN PHÁP";
        }
        // Nếu ở Tab "VIẾT MỚI", mặc định luôn là BIỆN PHÁP
        return "BIỆN PHÁP";
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

        // Tinh toan do dai tung muc dua tren gioi han trang chung
        let sectionLimitText = "";
        const pageLimitVal = document.getElementById('upgrade-page-limit')?.value || "15-20";
        const pages = pageLimitVal.split('-').map(p => parseInt(p.trim())).filter(p => !isNaN(p));
        if (pages.length > 0) {
            const idsList = ['1', '2', '2.3.1', '2.3.2', '2.3.3', '2.4', '3'];
            const maxPages = pages.length === 2 ? pages[1] : pages[0];
            const avgPages = (pages[0] + (pages[1] || pages[0])) / 2;

            const wordsPerPage = 300;
            const totalWordsTarget = Math.round(avgPages * wordsPerPage);
            const wordsPerSection = Math.round(totalWordsTarget / idsList.length);
            const hardLimit = wordsPerSection + 100;

            let strategy = "";
            if (maxPages <= 12) {
                strategy = "[AI CHIẾN LƯỢC VIẾT NGẮN/SÚC TÍCH]: Viết cực kỳ cô đọng.";
            } else {
                strategy = "[AI CHIẾN LƯỢC ĐẠT CHUẨN]: Phân tích sâu nếu thưa, viết gọn nếu đã dày.";
            }

            sectionLimitText = `\n\n[QUY ĐỊNH ĐỘ ĐÀI]:
            - ${strategy}
            - MỤC TIÊU: Khoảng ${wordsPerSection} từ.
            - GIỚI HẠN CỨNG: KHÔNG ĐƯỢC viết quá ${hardLimit} từ.`;
        }

        const originalHtml = btn.innerHTML;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG SOẠN...';
        btn.disabled = true;
        area.innerHTML = "<i>[Đang soạn thảo...]</i>";

        const docType = getDocType();
        const docName = docType === 'SKKN' ? "SÁNG KIẾN KINH NGHIỆM (SKKN)" : "BÁO CÁO BIỆN PHÁP GVG/GVCNG";

        let prompt = `Viết nội dung cho ${docName} kỳ thi ${roleText} với đề tài "${title}", mục ${id}.${sectionLimitText}`;
        if (id === '1') prompt = `Viết nội dung cho ${docName} kỳ thi ${roleText} đề tài "${title}", mục 1. ĐẶT VẤN ĐỀ (bao gồm 1.1 Lí do, 1.2 Mục đích, 1.3 Đối tượng).${sectionLimitText}`;
        if (id === '2') prompt = `Viết nội dung cho ${docName} kỳ thi ${roleText} đề tài "${title}", mục 2. THỰC TRẠNG VẤN ĐỀ (bao gồm 2.1 Cơ sở lý luận và 2.2 Cơ sở thực tiễn).${sectionLimitText}`;

        if (docType === 'SKKN') {
            prompt += `\nLưu ý quan trọng: Vì đây là SKKN, hãy sử dụng ngôn từ khoa học, mang tính chất nghiên cứu, đúc kết chuyên sâu.`;
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

    const writingProgressView = document.getElementById('writing-progress-view');
    const progressBarFill = document.getElementById('writing-progress-bar-fill');
    const progressPercent = document.getElementById('progress-percent');
    const progressStatus = document.getElementById('progress-status');
    const progressTitle = document.getElementById('progress-title');
    const progressLoadingIcon = document.getElementById('progress-loading-icon');
    const progressSuccessIcon = document.getElementById('progress-success-icon');
    const finishedActions = document.getElementById('finished-actions');

    async function writeAllSections(startIndex = 0, extraContext = "") {
        const title = document.getElementById('topic-title').value;
        const roleText = getRoleText();
        if (!title) return alert('Vui lòng nhập tên đề tài!');

        // Show progress modal
        if (writingProgressView) {
            writingProgressView.style.display = 'flex';
            progressLoadingIcon.style.display = 'block';
            progressSuccessIcon.style.display = 'none';
            finishedActions.style.display = 'none';
            progressTitle.innerText = "Đang nâng cấp biện pháp...";
            progressStatus.innerText = "Vui lòng đợi, AI đang xử lý nội dung từng phần.";
        }

        writeAllBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG VIẾT...';
        writeAllBtn.disabled = true;

        const ids = ['1', '2', '2.3.1', '2.3.2', '2.3.3', '2.4', '3'];

        for (let i = startIndex; i < ids.length; i++) {
            const id = ids[i];

            // Update UI Progress
            const percent = Math.round((i / ids.length) * 100);
            if (progressBarFill) progressBarFill.style.width = `${percent}%`;
            if (progressPercent) progressPercent.innerText = `${percent}%`;
            if (progressStatus) progressStatus.innerText = `Đang xử lý Mục ${id} (${i + 1}/${ids.length})...`;

            const area = document.querySelector(`.editor-area[data-id="${id}"]`);
            if (area) {
                area.innerHTML = "<i>[Đang soạn thảo...]</i>";
                // Tinh toan do dai tung muc dua tren gioi han trang
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

                let prompt = `Viết nội dung cho ${docName} kỳ thi ${roleText} với đề tài "${title}", mục ${id}.${sectionLimitText}`;
                if (id === '1') prompt = `Viết nội dung cho ${docName} kỳ thi ${roleText} đề tài "${title}", mục 1. ĐẶT VẤN ĐỀ (bao gồm 1.1 Lí do, 1.2 Mục đích, 1.3 Đối tượng).${sectionLimitText}`;
                if (id === '2') prompt = `Viết nội dung cho ${docName} kỳ thi ${roleText} đề tài "${title}", mục 2. THỰC TRẠNG VẤN ĐỀ (bao gồm 2.1 Cơ sở lý luận và 2.2 Cơ sở thực tiễn).${sectionLimitText}`;

                if (docType === 'SKKN') {
                    prompt += `\nLưu ý quan trọng: Vì đây là SKKN, hãy sử dụng ngôn từ khoa học, mang tính chất nghiên cứu, đúc kết chuyên sâu.`;
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
                        if (writingProgressView) writingProgressView.style.display = 'none';
                        return;
                    }
                }
            }
            await new Promise(r => setTimeout(r, 1200));
        }

        // Finished Successfully
        if (progressBarFill) progressBarFill.style.width = `100%`;
        if (progressPercent) progressPercent.innerText = `100%`;
        if (progressLoadingIcon) progressLoadingIcon.style.display = 'none';
        if (progressSuccessIcon) progressSuccessIcon.style.display = 'block';
        if (progressTitle) progressTitle.innerText = "Hoàn tất nâng cấp!";
        if (progressStatus) progressStatus.innerText = "Biện pháp của bạn đã được tối ưu hoàn chỉnh.";
        if (finishedActions) finishedActions.style.display = 'flex';

        writeAllBtn.innerHTML = '<i class="fas fa-magic"></i> VIẾT TOÀN BỘ (AI)';
        writeAllBtn.disabled = false;
    }

    document.getElementById('final-export-docx-btn')?.addEventListener('click', () => {
        exportWordBtn?.click();
    });

    document.getElementById('close-progress-btn')?.addEventListener('click', () => {
        if (writingProgressView) writingProgressView.style.display = 'none';
    });

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
                            new TextRun({ text: entry.label + " ", bold: true, size: 24, font: "Times New Roman" }),
                            new TextRun({ text: entry.value, size: 24, font: "Times New Roman" })
                        ]
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
                    // Add Title
                    children.push(new Paragraph({
                        children: [new TextRun({ text: task.title, bold: true, size: 26, font: "Times New Roman" })],
                        spacing: { before: 400, after: 200 }
                    }));

                    // Add Sub-title if exists
                    if (task.sub) {
                        children.push(new Paragraph({
                            children: [new TextRun({ text: task.sub, bold: true, size: 24, font: "Times New Roman" })],
                            spacing: { before: 200, after: 100 }
                        }));
                    }

                    const editor = document.querySelector(`.editor-area[data-id="${task.id}"]`);
                    const html = editor ? editor.innerHTML : " (Chưa nhập nội dung)";

                    // Simple HTML parsing for docx-js
                    const lines = html.split(/<br>|<div>|<\/div>/);

                    lines.forEach(line => {
                        const cleanLine = line.replace(/<[^>]*>/g, '').trim();
                        if (!cleanLine) return;

                        const boldParts = line.match(/<b>(.*?)<\/b>/g) || [];
                        const childrenRuns = [];

                        let remainingText = line;
                        if (boldParts.length > 0) {
                            boldParts.forEach(part => {
                                const [before, ...afterParts] = remainingText.split(part);
                                const boldText = part.replace(/<\/?b>/g, '');

                                if (before) childrenRuns.push(new TextRun({ text: before.replace(/<[^>]*>/g, ''), size: 24, font: "Times New Roman" }));
                                childrenRuns.push(new TextRun({ text: boldText, bold: true, size: 24, font: "Times New Roman" }));

                                remainingText = afterParts.join(part);
                            });
                            if (remainingText) childrenRuns.push(new TextRun({ text: remainingText.replace(/<[^>]*>/g, ''), size: 24, font: "Times New Roman" }));
                        } else {
                            childrenRuns.push(new TextRun({ text: cleanLine, size: 24, font: "Times New Roman" }));
                        }

                        children.push(new Paragraph({
                            alignment: AlignmentType.JUSTIFY,
                            indent: { firstLine: 454 },
                            spacing: { line: 360, after: 120 },
                            children: childrenRuns
                        }));
                    });
                });

                const doc = new Document({
                    sections: [{
                        properties: { page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1701 } } },
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
                if (topicInput) topicInput.value = suggestions[0].replace(/["']/g, '').trim();

                // Show other suggestions via prompt
                setTimeout(() => {
                    const choiceList = suggestions.map((s, i) => `${i + 1}. ${s.trim()}`).join('\n');
                    const choice = prompt(`AI đã chọn câu hay nhất. Xem các lựa chọn khác:\n\n${choiceList}\n\nNhập số 1-5 để chọn lại:`);
                    if (choice && !isNaN(choice) && suggestions[parseInt(choice) - 1]) {
                        topicInput.value = suggestions[parseInt(choice) - 1].replace(/["']/g, '').trim();
                    }
                }, 500);
            }
        }

        aiSuggestBtn.innerHTML = originalHtml;
        aiSuggestBtn.disabled = false;
    });

    const completeBtn = document.getElementById('complete-btn');
    completeBtn?.addEventListener('click', () => {
        alert("Chúc mừng! Bạn đã hoàn thành báo cáo biện pháp GVG/GVCNG. Hãy nhấn Xuất Word để tải file về.");
        completeBtn.innerHTML = '<i class="fas fa-check"></i> ĐÃ HOÀN THÀNH';
        completeBtn.style.opacity = '0.8';
    });

    // Appraisal Logic
    const selectFileBtn = document.getElementById('select-file-btn');
    const fileInput = document.getElementById('file-upload-input');
    const uploadDropzone = document.getElementById('upload-dropzone');
    const uploadView = document.getElementById('appraisal-upload-view');
    const resultsView = document.getElementById('appraisal-results-view');
    const reUploadBtn = document.getElementById('re-upload-btn');

    selectFileBtn?.addEventListener('click', () => fileInput.click());

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

    uploadDropzone?.addEventListener('drop', (e) => {
        e.preventDefault();
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

    let lastAppraisalData = null;
    let lastAppraisedText = "";

    async function startAnalysis(file) {
        const originalContent = uploadDropzone.innerHTML;
        uploadDropzone.innerHTML = `
            <div class="upload-icon"><i class="fas fa-spinner fa-spin"></i></div>
            <h3>Đang thẩm định chuyên sâu "${file.name}"...</h3>
            <p>AI đang phân tích phong cách hành văn, cấu trúc và đặc trưng sư phạm. Vui lòng đợi.</p>
        `;

        try {
            const appraisalType = document.querySelector('input[name="appraisal_type"]:checked')?.value || "BIỆN PHÁP";
            const fileText = await extractText(file);
            if (!fileText || fileText.length < 100) {
                alert("File không chứa đủ văn bản hoặc định dạng không được hỗ trợ.");
                uploadDropzone.innerHTML = originalContent;
                return;
            }

            lastAppraisedText = fileText; // Store for topic analysis

            const prompt = `Bạn là chuyên gia thẩm định ${appraisalType} với 20 năm kinh nghiệm. 
            Nếu là BIỆN PHÁP: Tập trung vào tính thực tiễn, quy mô lớp học và hiệu quả giảng dạy trực tiếp.
            Nếu là SÁNG KIẾN KINH NGHIỆM (SKKN): Tập trung vào cơ sở lý luận, tính mới, tính khoa học và khả năng nhân rộng ngoài phạm vi một lớp học/trường học.

            Hãy ĐỌC KỸ và PHÂN TÍCH nội dung thực tế từ file giáo viên tải lên dưới đây để thẩm định đúng tiêu chuẩn của một bài ${appraisalType}:
            
            """${fileText.substring(0, 15000)}"""

            YÊU CẦU QUAN TRỌNG:
            1. Tuyệt đối KHÔNG sử dụng lại nội dung mẫu hay từ hình ảnh hướng dẫn.
            2. Phải phân tích dựa trên NGỮ CẢNH THỰC của file (môn học nào, khối lớp nào, giải pháp là gì).
            3. "patterns": Hãy chỉ ra các đặc điểm hành văn THỰC TẾ trong bài này khiến AI nghi ngờ (Vd: cấu trúc câu quá cân đối, dùng từ 'kỷ nguyên số' quá đà, viết quá khách quan thiếu cảm xúc...).
            4. "suggestions": Đưa ra các giải pháp cải thiện THẬT SỰ cho bài này (Vd: bài đang thiếu ví dụ ở phần giải pháp 2, cần thêm số liệu khảo sát học sinh lớp này...).

            Trả về kết quả JSON chính xác:
            {
              "summary": "Tóm tắt đánh giá cụ thể cho bài này",
              "quality_score": 0-100,
              "plagiarism_risk": 0-100,
              "ai_risk": 0-100,
              "status_label": "Giỏi/Khá/Trung bình/Yếu",
              "structure": { "toc": bool, "intro": bool, "reality": bool, "solution": bool, "conclusion": bool },
              "criteria": { "scientific": 0-10, "practical": 0-10, "innovation": 0-10, "applied_ability": 0-10, "demonstrated_effect": 0-10, "language": 0-10 },
              "plagiarism_examples": [ { "quote": "...", "source": "..." } ],
              "ai_details": {
                "perplexity": 20-80,
                "burstiness": 20-80,
                "pattern_score": 0-100,
                "patterns": ["Đặc điểm hành văn thực tế bài này 1", "Đặc điểm 2..."],
                "suggestions": ["Gợi ý cải thiện thực tế bài này 1", "Gợi ý 2..."]
              }
            }`;

            const result = await callGemini(prompt, () => startAnalysis(file));

            if (result.text) {
                let jsonText = result.text;
                if (jsonText.includes('```json')) jsonText = jsonText.split('```json')[1].split('```')[0];
                else if (jsonText.includes('```')) jsonText = jsonText.split('```')[1].split('```')[0];

                const data = JSON.parse(jsonText.trim());
                lastAppraisalData = data; // Store for export
                updateAppraisalUI(data);

                uploadView.style.display = 'none';
                resultsView.style.display = 'block';
            } else {
                throw new Error(result.error || "Không nhận được phản hồi từ AI");
            }

        } catch (error) {
            console.error("Appraisal Error:", error);
            alert("Lỗi thẩm định: " + error.message);
        } finally {
            uploadDropzone.innerHTML = originalContent;
        }
    }

    // Export Appraisal Report to DOCX
    document.getElementById('export-appraisal-btn')?.addEventListener('click', async () => {
        if (!lastAppraisalData) return;

        try {
            const doc = new docx.Document({
                styles: {
                    default: {
                        paragraph: {
                            run: {
                                size: 24,
                                font: "Times New Roman"
                            }
                        }
                    }
                },
                sections: [{
                    properties: {},
                    children: [
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: "BÁO CÁO THẨM ĐỊNH BIỆN PHÁP + SKKN", bold: true, size: 28 })],
                            alignment: docx.AlignmentType.CENTER,
                        }),
                        new docx.Paragraph({ text: "" }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: "Tóm tắt đánh giá:", bold: true })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: lastAppraisalData.summary })],
                        }),
                        new docx.Paragraph({ text: "" }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: "Chỉ số chất lượng:", bold: true })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Điểm chất lượng tổng thể: ${lastAppraisalData.quality_score}/100 (${lastAppraisalData.status_label})` })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Nguy cơ đạo văn: ${lastAppraisalData.plagiarism_risk}%` })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Nguy cơ AI: ${lastAppraisalData.ai_risk}%` })],
                        }),
                        new docx.Paragraph({ text: "" }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: "Chi tiết điểm số (thang điểm 10):", bold: true })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Tính khoa học: ${lastAppraisalData.criteria.scientific}/10` })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Tính thực tiễn: ${lastAppraisalData.criteria.practical}/10` })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Tính mới/sáng tạo: ${lastAppraisalData.criteria.innovation}/10` })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Khả năng áp dụng: ${lastAppraisalData.criteria.applied_ability}/10` })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Hiệu quả minh chứng: ${lastAppraisalData.criteria.demonstrated_effect}/10` })],
                        }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: `- Ngôn ngữ và trình bày: ${lastAppraisalData.criteria.language}/10` })],
                        }),
                        new docx.Paragraph({ text: "" }),
                        new docx.Paragraph({
                            children: [new docx.TextRun({ text: "Gợi ý cải thiện:", bold: true })],
                        }),
                        ...lastAppraisalData.ai_details.suggestions.map(s => new docx.Paragraph({
                            children: [new docx.TextRun({ text: `• ${s}` })],
                            bullet: { level: 0 }
                        })),
                    ],
                }],
            });

            const blob = await docx.Packer.toBlob(doc);
            saveAs(blob, "Bao_cao_Tham_dinh.docx");
        } catch (error) {
            console.error("Export Error:", error);
            alert("Lỗi khi xuất file Word.");
        }
    });

    let selectedTitle = "";
    let isNewTitle = false;

    function showUpgradePlan(title, isNew) {
        selectedTitle = title;
        isNewTitle = isNew;

        const upgradeStepsList = document.getElementById('upgrade-steps-list');
        const alertAiScore = document.getElementById('alert-ai-score');
        const upgradePlanView = document.getElementById('upgrade-plan-view');

        // Populate steps from appraisal suggestions
        if (lastAppraisalData && lastAppraisalData.ai_details.suggestions) {
            upgradeStepsList.innerHTML = lastAppraisalData.ai_details.suggestions.map(s => `<li>${s}</li>`).join('');
            alertAiScore.innerText = lastAppraisalData.ai_risk;
        } else {
            upgradeStepsList.innerHTML = '<li>Hoàn thiện các nội dung còn thiếu trong báo cáo.</li><li>Nâng cấp ngôn ngữ chuyên môn theo chuẩn GVG.</li>';
            alertAiScore.innerText = "50";
        }

        topicAnalysisView.style.display = 'none';
        upgradePlanView.style.display = 'flex';
    }

    document.getElementById('keep-old-topic-btn')?.addEventListener('click', () => {
        const currentTitle = document.getElementById('topic-title')?.value || "Đề tài hiện tại";
        showUpgradePlan(currentTitle, false);
    });

    document.getElementById('close-upgrade-plan')?.addEventListener('click', () => {
        document.getElementById('upgrade-plan-view').style.display = 'none';
    });

    // Mini export buttons inside upgrade plan
    document.querySelector('#upgrade-plan-view .mini-export-btn.docx')?.addEventListener('click', () => {
        document.getElementById('export-appraisal-btn')?.click();
    });

    document.getElementById('confirm-upgrade-btn')?.addEventListener('click', () => {
        // Set the title
        const topicInput = document.getElementById('topic-title');
        if (topicInput) topicInput.value = selectedTitle;

        // Get page limit
        const pageLimit = document.getElementById('upgrade-page-limit')?.value || "15-20";

        // Close view and switch to sections
        document.getElementById('upgrade-plan-view').style.display = 'none';

        // Find the first relevant section to switch to
        switchSection('section1');
        currentSectionIndex = sectionOrder.indexOf('section1');

        // Update nav UI
        navItems.forEach(l => l.classList.remove('active'));
        document.querySelector('[data-section="section1"]')?.classList.add('active');

        // Prepare context for AI writing
        const improvements = lastAppraisalData?.ai_details.suggestions.join("\n- ") || "";
        const appraisalType = document.querySelector('input[name="appraisal_type"]:checked')?.value || "BIỆN PHÁP";

        let contextPrompt = isNewTitle
            ? `Hãy VIẾT LẠI HOÀN CHỈNH ${appraisalType} dựa trên TÊN ĐỀ TÀI MỚI: "${selectedTitle}". Cần khắc phục triệt để các hạn chế sau: \n- ${improvements}`
            : `HÃY HOÀN THIỆN các phần còn thiếu và sửa đổi những hạn chế của ${appraisalType}: "${selectedTitle}". Tập trung vào: \n- ${improvements}`;

        if (appraisalType === 'SKKN') {
            contextPrompt += `\nLưu ý: Vì đây là SÁNG KIẾN KINH NGHIỆM, hãy chú trọng vào tính khoa học, tính mới và đúc kết thành bài học kinh nghiệm sâu sắc.`;
        } else {
            contextPrompt += `\nLưu ý: Vì đây là BIỆN PHÁP GVG/GVCNG, hãy chú trọng vào tính thực tiễn, các bước triển khai cụ thể tại lớp và minh chứng hiệu quả trên học sinh.`;
        }

        // Giới hạn độ dài sẽ do writeAllSections tự động tính toán từ input '#upgrade-page-limit'
        contextPrompt += `\n\n[QUY ĐỊNH VỀ TRÌNH BÀY]: Hãy tuân thủ nghiêm ngặt số lượng trang đã được thiết lập. 
        Mục tiêu cuối cùng là bản in Word phải đẹp, chuyên nghiệp và đúng độ dài ${pageLimit} trang.`;

        // Trigger automatic writing of all sections with this context
        writeAllSections(0, contextPrompt);
    });

    // Topic Analysis (Image A3) Logic
    const refineTopicBtn = document.getElementById('refine-topic-btn');
    const topicAnalysisView = document.getElementById('topic-analysis-view');
    const closeTopicAnalysis = document.getElementById('close-topic-analysis');
    const topicSuggestionsList = document.getElementById('topic-suggestions-list');

    refineTopicBtn?.addEventListener('click', async () => {
        if (!lastAppraisedText) {
            alert('Vui lòng thẩm định file trước khi sử dụng tính năng này.');
            return;
        }

        topicAnalysisView.style.display = 'flex';
        topicSuggestionsList.innerHTML = '<div class="loading-suggestions"><i class="fas fa-spinner fa-spin"></i> Đang phân tích đề tài tối ưu từ bài viết của bạn...</div>';

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
                topicSuggestionsList.innerHTML = `
                    <div class="loading-suggestions" style="color: #e74c3c;">
                        <i class="fas fa-exclamation-triangle"></i><br><br>
                        ${result.error === 'Chưa cấu hình API Key.' ? 'Chưa có API Key. Vui lòng cấu hình trong menu.' : 'Lỗi: ' + result.error}
                    </div>`;
                return;
            }

            if (result.text) {
                let jsonText = result.text;
                if (jsonText.includes('```json')) jsonText = jsonText.split('```json')[1].split('```')[0];
                else if (jsonText.includes('```')) jsonText = jsonText.split('```')[1].split('```')[0];

                try {
                    const data = JSON.parse(jsonText.trim());
                    if (!data.suggestions || data.suggestions.length === 0) throw new Error("No suggestions");

                    topicSuggestionsList.innerHTML = data.suggestions.map(s => `
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
                            showUpgradePlan(title, true);
                        });
                    });
                } catch (e) {
                    console.error("JSON Parse Error:", e, jsonText);
                    topicSuggestionsList.innerHTML = '<div class="loading-suggestions">AI trả về dữ liệu không đúng cấu trúc. Vui lòng thử lại.</div>';
                }
            } else {
                topicSuggestionsList.innerHTML = '<div class="loading-suggestions">Không nhận được phản hồi từ AI.</div>';
            }
        } catch (error) {
            console.error("Refine Topic Error:", error);
            topicSuggestionsList.innerHTML = `<div class="loading-suggestions">Lỗi hệ thống: ${error.message}</div>`;
        }
    });

    closeTopicAnalysis?.addEventListener('click', () => {
        topicAnalysisView.style.display = 'none';
    });

    function updateAppraisalUI(data) {
        // Summary
        document.getElementById('analysis-summary-text').innerText = data.summary;

        // Overall Score
        updateCircularProgress('CHẤT LƯỢNG TỔNG THỂ', data.quality_score, data.status_label);
        updateCircularProgress('NGUY CƠ ĐẠO VĂN', data.plagiarism_risk);
        updateCircularProgress('NGUY CƠ AI', data.ai_risk);

        // Structure
        const structureList = document.querySelector('.structure-list');
        structureList.innerHTML = `
            <li class="${data.structure.toc ? 'check' : 'error'}"><i class="fas ${data.structure.toc ? 'fa-check-circle' : 'fa-times-circle'}"></i> Mục lục</li>
            <li class="${data.structure.intro ? 'check' : 'error'}"><i class="fas ${data.structure.intro ? 'fa-check-circle' : 'fa-times-circle'}"></i> 1. Đặt vấn đề</li>
            <li class="${data.structure.reality ? 'check' : 'error'}"><i class="fas ${data.structure.reality ? 'fa-check-circle' : 'fa-times-circle'}"></i> 2. Thực trạng vấn đề</li>
            <li class="${data.structure.solution ? 'check' : 'warn'}"><i class="fas ${data.structure.solution ? 'fa-check-circle' : 'fa-exclamation-triangle'}"></i> 3. Nội dung giải pháp</li>
            <li class="${data.structure.conclusion ? 'check' : 'error'}"><i class="fas ${data.structure.conclusion ? 'fa-check-circle' : 'fa-times-circle'}"></i> 4. Kết luận và kiến nghị</li>
        `;

        // Bar Chart
        updateBarChart(data.criteria);

        // Radar Chart
        drawRadarChart(Object.values(data.criteria).map(v => v / 10));

        // Plagiarism
        const plagiarismList = document.querySelector('.plagiarism-list');
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

        // AI Details - Strictly matching screenshot labels
        const aiStats = document.querySelector('.ai-stats-grid');
        aiStats.innerHTML = `
            <div class="ai-stat-item">
                <div class="stat-value">${data.ai_details.perplexity}</div>
                <div class="stat-label">PERPLEXITY</div>
                <div class="stat-info">Thấp = Đậm chất AI</div>
            </div>
            <div class="ai-stat-item">
                <div class="stat-value">${data.ai_details.burstiness}</div>
                <div class="stat-label">BURSTINESS</div>
                <div class="stat-info">Cao = Đậm chất Người</div>
            </div>
            <div class="ai-stat-item">
                <div class="stat-value">${data.ai_details.pattern_score}</div>
                <div class="stat-label">PATTERN</div>
                <div class="stat-info">Tỷ lệ khớp mẫu câu AI</div>
            </div>
        `;

        const aiPatternsTitle = document.querySelector('.ai-patterns h4');
        if (aiPatternsTitle) aiPatternsTitle.innerText = 'Mẫu câu AI phát hiện:';

        const aiPatternsList = document.querySelector('.ai-patterns ul');
        aiPatternsList.innerHTML = data.ai_details.patterns.map(p => `<li>${p}</li>`).join('');

        const aiSuggestionsTitle = document.querySelector('.ai-suggestions h4');
        if (aiSuggestionsTitle) aiSuggestionsTitle.innerHTML = '<i class="far fa-lightbulb"></i> Gợi ý giảm AI:';

        const aiSuggestionsList = document.querySelector('.ai-suggestions ul');
        // Checkmark icons are handled by CSS ::before in style.css
        aiSuggestionsList.innerHTML = data.ai_details.suggestions.map(s => `<li>${s}</li>`).join('');
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
        if (title.includes('CHẤT LƯỢNG')) valueBox.innerHTML = `${value}<span>/100 điểm</span>`;
        else valueBox.innerHTML = `${value}%<span>${title.includes('ĐẠO VĂN') ? 'trùng lặp' : 'nguy cơ AI'}</span>`;

        // Update circle
        const offset = circumference - (value / 100) * circumference;
        circle.style.strokeDasharray = `${circumference}`;
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
            badge.className = `status-badge ${status === 'Giỏi' ? 'success' : status === 'Khá' ? 'warning' : 'danger'}`;
        }
    }

    function updateBarChart(criteria) {
        const barItems = document.querySelectorAll('.bar-item');
        const mapping = {
            'Tính khoa học': criteria.scientific,
            'Tính thực tiễn': criteria.practical,
            'Tính mới': criteria.innovation,
            'Khả năng áp dụng': criteria.applied_ability,
            'Hiệu quả': criteria.demonstrated_effect,
            'Ngôn ngữ': criteria.language
        };

        barItems.forEach(item => {
            const label = item.querySelector('.bar-label').innerText.trim().toLowerCase();
            let val = 0;

            if (label.includes('khoa học')) val = criteria.scientific;
            else if (label.includes('thực tiễn')) val = criteria.practical;
            else if (label.includes('mới')) val = criteria.innovation;
            else if (label.includes('áp dụng')) val = criteria.applied_ability;
            else if (label.includes('hiệu quả')) val = criteria.demonstrated_effect;
            else if (label.includes('ngôn ngữ')) val = criteria.language;

            item.querySelector('.bar-fill').style.width = `${val * 10}%`;
            item.querySelector('.bar-value').innerText = `${val}/10`;
            item.querySelector('.bar-fill').className = `bar-fill ${val >= 8 ? 'success' : val >= 6 ? 'warning' : 'danger'}`;
        });
    }

    reUploadBtn?.addEventListener('click', (e) => {
        e.preventDefault();
        resultsView.style.display = 'none';
        uploadView.style.display = 'block';
        fileInput.value = '';
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

    infoForm?.addEventListener('submit', (e) => { e.preventDefault(); switchSection('outline'); });
});
