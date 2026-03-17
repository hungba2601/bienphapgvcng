document.addEventListener('DOMContentLoaded', () => {
    // Selectors
    const navItems = document.querySelectorAll('.nav-item');
    const sections = document.querySelectorAll('.content-section');
    const infoForm = document.getElementById('info-form');
    const aiSuggestBtn = document.querySelector('.ai-btn');
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

        const originalHtml = btn.innerHTML;
        btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG SOẠN...';
        btn.disabled = true;
        area.innerHTML = "<i>[Đang soạn thảo...]</i>";

        let prompt = `Viết báo cáo biện pháp cho kỳ thi ${roleText} với đề tài "${title}", mục ${id}.`;
        if (id === '1') prompt = `Viết báo cáo biện pháp kỳ thi ${roleText} đề tài "${title}", mục 1. ĐẶT VẤN ĐỀ (bao gồm 1.1 Lí do, 1.2 Mục đích, 1.3 Đối tượng).`;
        if (id === '2') prompt = `Viết báo cáo biện pháp kỳ thi ${roleText} đề tài "${title}", mục 2. THỰC TRẠNG VẤN ĐỀ (bao gồm 2.1 Cơ sở lý luận và 2.2 Cơ sở thực tiễn).`;

        if (userPrompt) {
            prompt += `\nLưu ý: Viết dựa trên gợi ý sau của người dùng: "${userPrompt}"`;
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

    async function writeAllSections(startIndex = 0) {
        const title = document.getElementById('topic-title').value;
        if (!title) return alert('Vui lòng nhập tên đề tài!');

        writeAllBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> ĐANG VIẾT...';
        writeAllBtn.disabled = true;

        const ids = ['1', '2', '2.3.1', '2.3.2', '2.3.3', '2.4', '3'];

        for (let i = startIndex; i < ids.length; i++) {
            const id = ids[i];
            const area = document.querySelector(`.editor-area[data-id="${id}"]`);
            if (area) {
                area.innerHTML = "<i>[Đang soạn thảo...]</i>";
                let prompt = `Viết báo cáo biện pháp GVG đề tài "${title}", mục ${id}.`;
                if (id === '1') prompt = `Viết báo cáo biện pháp GVG đề tài "${title}", mục 1. ĐẶT VẤN ĐỀ (bao gồm 1.1 Lí do, 1.2 Mục đích, 1.3 Đối tượng).`;
                if (id === '2') prompt = `Viết báo cáo biện pháp GVG đề tài "${title}", mục 2. THỰC TRẠNG VẤN ĐỀ (bao gồm 2.1 Cơ sở lý luận và 2.2 Cơ sở thực tiễn).`;

                const result = await callGemini(prompt, () => writeAllSections(i));
                if (result.text) {
                    area.innerHTML = formatAiResponse(result.text);
                } else {
                    area.innerHTML = `<span style="color:red">[LỖI: ${result.error}]</span>`;
                    if (result.error.includes('Hết hạn mức')) {
                        writeAllBtn.innerHTML = '<i class="fas fa-magic"></i> TIẾP TỤC VIẾT (AI)';
                        writeAllBtn.disabled = false;
                        return;
                    }
                }
            }
            await new Promise(r => setTimeout(r, 1200));
        }

        writeAllBtn.innerHTML = '<i class="fas fa-magic"></i> VIẾT TOÀN BỘ (AI)';
        writeAllBtn.disabled = false;
    }

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
                    children: [new TextRun({ text: school.toUpperCase(), size: 28, font: "Times New Roman" })]
                }));

                children.push(new Paragraph({ spacing: { before: 2000 } })); // Spacer

                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "BÁO CÁO BIỆN PHÁP", bold: true, size: 36, font: "Times New Roman" })]
                }));

                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "GÓP PHẦN NÂNG CAO CHẤT LƯỢNG CÔNG TÁC GIẢNG DẠY", size: 24, font: "Times New Roman" })]
                }));

                children.push(new Paragraph({ spacing: { before: 800 } })); // Spacer

                children.push(new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: "ĐỀ TÀI:", bold: true, size: 28, font: "Times New Roman" })]
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
                            new TextRun({ text: entry.label + " ", bold: true, size: 26, font: "Times New Roman" }),
                            new TextRun({ text: entry.value, size: 26, font: "Times New Roman" })
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
                        children: [new TextRun({ text: task.title, bold: true, size: 28, font: "Times New Roman" })],
                        spacing: { before: 400, after: 200 }
                    }));

                    // Add Sub-title if exists
                    if (task.sub) {
                        children.push(new Paragraph({
                            children: [new TextRun({ text: task.sub, bold: true, size: 26, font: "Times New Roman" })],
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

                                if (before) childrenRuns.push(new TextRun({ text: before.replace(/<[^>]*>/g, ''), size: 26, font: "Times New Roman" }));
                                childrenRuns.push(new TextRun({ text: boldText, bold: true, size: 26, font: "Times New Roman" }));

                                remainingText = afterParts.join(part);
                            });
                            if (remainingText) childrenRuns.push(new TextRun({ text: remainingText.replace(/<[^>]*>/g, ''), size: 26, font: "Times New Roman" }));
                        } else {
                            childrenRuns.push(new TextRun({ text: cleanLine, size: 26, font: "Times New Roman" }));
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

    aiSuggestBtn?.addEventListener('click', async () => {
        const topicInput = document.getElementById('topic-title');
        const roleText = getRoleText();
        const subject = document.getElementById('subject-name')?.value || "một môn học";
        const grade = document.getElementById('grade-level')?.value || "các cấp";
        const className = document.getElementById('class-name')?.value || "";

        const originalHtml = aiSuggestBtn.innerHTML;
        aiSuggestBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Đang nghĩ...';
        aiSuggestBtn.disabled = true;

        const prompt = `Đề xuất 1 tên đề tài biện pháp cho kỳ thi ${roleText} hay, sáng tạo, thực tiễn cho môn ${subject}, ${grade} ${className}. 
        Yêu cầu: Viết nội dung tên đề tài duy nhất, không thêm giải thích, không để trong ngoặc kép. Tên đề tài nên bắt đầu bằng các cụm như "Một số biện pháp...", "Ứng dụng...", "Giải pháp nâng cao..."`;

        const result = await callGemini(prompt, () => aiSuggestBtn.click());

        if (result.text) {
            if (topicInput) topicInput.value = result.text.replace(/["']/g, ''); // Remove quotes if AI added them
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

    infoForm?.addEventListener('submit', (e) => { e.preventDefault(); switchSection('outline'); });
});
