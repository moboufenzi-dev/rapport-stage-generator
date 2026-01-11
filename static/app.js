// ===== STATE =====
let chapters = [];
let glossary = [];
let figures = [];
let ganttTasks = [];
let logos = {
    logo_ecole: null,
    logo_entreprise: null,
    image_centrale: null
};

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', () => {
    loadFromLocalStorage();
    initAutoSave();
    initDefaultChapters();
    initTabs();
    updatePreview();

    // Auto-update preview on any change
    document.querySelectorAll('input, select, textarea').forEach(el => {
        el.addEventListener('input', debounce(updatePreview, 300));
        el.addEventListener('change', updatePreview);
    });

    // Initialize Gantt section visibility
    toggleGanttSection();
});

function initTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const tabName = btn.getAttribute('data-tab');
            goToTab(tabName);
        });
    });
}

// ===== TABS =====
function goToTab(tabName) {
    // Hide all tab contents
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    // Deactivate all tab buttons
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

    // Show selected tab content
    const content = document.getElementById('tab-' + tabName);
    if (content) content.classList.add('active');

    // Activate selected tab button
    const btn = document.querySelector(`.tab-btn[data-tab="${tabName}"]`);
    if (btn) btn.classList.add('active');
}

// ===== GANTT TOGGLE =====
function toggleGanttSection() {
    const checkbox = document.getElementById('include_gantt');
    const section = document.getElementById('ganttSection');
    if (checkbox && section) {
        if (checkbox.checked) {
            section.classList.remove('hidden');
        } else {
            section.classList.add('hidden');
        }
    }
    updatePreview();
}

// ===== CHAPTERS =====
function initDefaultChapters() {
    if (chapters.length === 0) {
        chapters = [
            { id: 1, title: 'Introduction', level: 1, children: [] },
            { id: 2, title: 'Présentation de l\'entreprise', level: 1, children: [
                { id: 21, title: 'Histoire et activités', level: 2, children: [] },
                { id: 22, title: 'Organisation', level: 2, children: [] }
            ]},
            { id: 3, title: 'Missions et objectifs', level: 1, children: [] },
            { id: 4, title: 'Travail réalisé', level: 1, children: [] },
            { id: 5, title: 'Bilan', level: 1, children: [] },
            { id: 6, title: 'Conclusion', level: 1, children: [] }
        ];
    }
    renderChapters();
}

function renderChapters() {
    const container = document.getElementById('chaptersContainer');
    container.innerHTML = '';

    let chapterNum = 0;
    chapters.forEach((chapter, index) => {
        chapterNum++;
        container.appendChild(createChapterElement(chapter, index, chapterNum.toString()));

        chapter.children.forEach((sub, subIndex) => {
            container.appendChild(createChapterElement(sub, subIndex, `${chapterNum}.${subIndex + 1}`, index));
        });
    });

    saveToLocalStorage();
    updatePreview();
}

function createChapterElement(chapter, index, number, parentIndex = null) {
    const div = document.createElement('div');
    div.className = `chapter-item level-${chapter.level}`;

    const canAddSub = chapter.level < 2;
    const canMove = parentIndex === null;

    div.innerHTML = `
        <span class="chapter-number">${number}</span>
        <input type="text" class="chapter-input" value="${chapter.title}"
               onchange="updateChapterTitle(${chapter.id}, this.value)">
        <div class="chapter-actions">
            ${canAddSub ? `<button class="add-sub" onclick="addSubChapter(${chapter.id})">+</button>` : ''}
            ${canMove ? `<button onclick="moveChapter(${index}, -1)">↑</button>` : ''}
            ${canMove ? `<button onclick="moveChapter(${index}, 1)">↓</button>` : ''}
            <button class="delete" onclick="deleteChapter(${chapter.id})">×</button>
        </div>
    `;

    return div;
}

function addChapter() {
    chapters.push({ id: Date.now(), title: 'Nouveau chapitre', level: 1, children: [] });
    renderChapters();
}

function addSubChapter(parentId) {
    for (let ch of chapters) {
        if (ch.id === parentId) {
            ch.children.push({ id: Date.now(), title: 'Nouvelle section', level: 2, children: [] });
            break;
        }
    }
    renderChapters();
}

function updateChapterTitle(id, title) {
    function find(items) {
        for (let item of items) {
            if (item.id === id) { item.title = title; return; }
            if (item.children) find(item.children);
        }
    }
    find(chapters);
    saveToLocalStorage();
    updatePreview();
}

function deleteChapter(id) {
    function del(items) {
        for (let i = 0; i < items.length; i++) {
            if (items[i].id === id) { items.splice(i, 1); return true; }
            if (items[i].children && del(items[i].children)) return true;
        }
        return false;
    }
    del(chapters);
    renderChapters();
}

function moveChapter(index, direction) {
    const newIndex = index + direction;
    if (newIndex < 0 || newIndex >= chapters.length) return;
    [chapters[index], chapters[newIndex]] = [chapters[newIndex], chapters[index]];
    renderChapters();
}

// ===== GLOSSARY =====
function addGlossaryItem() {
    const term = document.getElementById('glossaryTerm').value.trim();
    const def = document.getElementById('glossaryDef').value.trim();
    if (!term) return;

    glossary.push({ term, definition: def || '[Définition]' });
    document.getElementById('glossaryTerm').value = '';
    document.getElementById('glossaryDef').value = '';
    renderGlossary();
}

function renderGlossary() {
    const container = document.getElementById('glossaryContainer');
    container.innerHTML = glossary.map((item, i) => `
        <div class="list-item">
            <span class="list-item-num">${i + 1}.</span>
            <span class="list-item-content"><strong>${item.term}</strong> : ${item.definition}</span>
            <button class="list-item-delete" onclick="deleteGlossaryItem(${i})">×</button>
        </div>
    `).join('');
    saveToLocalStorage();
    updatePreview();
}

function deleteGlossaryItem(index) {
    glossary.splice(index, 1);
    renderGlossary();
}

// ===== FIGURES =====
function addFigure() {
    const name = document.getElementById('figureName').value.trim();
    const page = document.getElementById('figurePage').value.trim();
    if (!name) return;

    figures.push({ name, page: page || '-' });
    document.getElementById('figureName').value = '';
    document.getElementById('figurePage').value = '';
    renderFigures();
}

function renderFigures() {
    const container = document.getElementById('figuresContainer');
    container.innerHTML = figures.map((item, i) => `
        <div class="list-item">
            <span class="list-item-num">Fig. ${i + 1}</span>
            <span class="list-item-content">${item.name}</span>
            <span style="color:#6b7280;font-size:0.75rem;">p.${item.page}</span>
            <button class="list-item-delete" onclick="deleteFigure(${i})">×</button>
        </div>
    `).join('');
    saveToLocalStorage();
    updatePreview();
}

function deleteFigure(index) {
    figures.splice(index, 1);
    renderFigures();
}

// ===== GANTT =====
function addGanttTask() {
    const task = document.getElementById('ganttTask').value.trim();
    const start = document.getElementById('ganttStart').value;
    const end = document.getElementById('ganttEnd').value;
    if (!task || !start || !end) return;

    ganttTasks.push({ task, start, end });
    document.getElementById('ganttTask').value = '';
    document.getElementById('ganttStart').value = '';
    document.getElementById('ganttEnd').value = '';
    renderGantt();
}

function renderGantt() {
    const container = document.getElementById('ganttContainer');
    container.innerHTML = ganttTasks.map((item, i) => `
        <div class="list-item">
            <span class="list-item-num">${i + 1}.</span>
            <span class="list-item-content">${item.task}</span>
            <span style="color:#6b7280;font-size:0.7rem;">${formatDateShort(item.start)} → ${formatDateShort(item.end)}</span>
            <button class="list-item-delete" onclick="deleteGanttTask(${i})">×</button>
        </div>
    `).join('');
    saveToLocalStorage();
    updatePreview();
}

function deleteGanttTask(index) {
    ganttTasks.splice(index, 1);
    renderGantt();
}

function formatDateShort(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return `${d.getDate()}/${d.getMonth() + 1}`;
}

// ===== LOGO PREVIEW =====
function previewLogo(input, previewId, logoKey) {
    const preview = document.getElementById(previewId);
    const file = input.files[0];

    if (file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            const base64 = e.target.result;
            preview.innerHTML = `<img src="${base64}" alt="Logo">`;
            logos[logoKey] = base64;
            saveToLocalStorage();
            updatePreview();
        };
        reader.readAsDataURL(file);
    }
}

// ===== LOCAL STORAGE =====
function saveToLocalStorage() {
    const data = collectFormData();
    localStorage.setItem('reportDataV3', JSON.stringify(data));
    showSaveIndicator();
}

function loadFromLocalStorage() {
    const saved = localStorage.getItem('reportDataV3');
    if (!saved) return;

    try {
        const data = JSON.parse(saved);

        // Restore form fields
        const fields = [
            'prenom', 'nom', 'formation', 'ecole', 'annee_scolaire',
            'entreprise_nom', 'entreprise_secteur', 'entreprise_ville',
            'tuteur_nom', 'tuteur_poste', 'tuteur_academique_nom', 'tuteur_academique_poste',
            'date_debut', 'date_fin', 'poste', 'mission_principale',
            'font_family', 'font_size', 'line_spacing',
            'title1_size', 'title1_color', 'title2_size', 'title2_color',
            'title3_size', 'title3_color', 'margin_top', 'margin_bottom',
            'margin_left', 'margin_right'
        ];

        fields.forEach(field => {
            const el = document.getElementById(field);
            if (el && data[field]) el.value = data[field];
        });

        // Restore checkboxes
        const checkboxes = [
            'include_cover', 'include_thanks', 'include_toc', 'include_figures_list',
            'include_abstract', 'include_glossary', 'include_gantt', 'include_annexes',
            'title1_bold', 'title2_bold', 'title3_italic',
            'show_page_number', 'show_student_name'
        ];

        checkboxes.forEach(field => {
            const el = document.getElementById(field);
            if (el && data[field] !== undefined) el.checked = data[field];
        });

        // Restore radio buttons
        if (data.cover_model) {
            const radio = document.querySelector(`input[name="cover_model"][value="${data.cover_model}"]`);
            if (radio) radio.checked = true;
        }

        // Restore arrays
        if (data.chapters) chapters = data.chapters;
        if (data.glossary) glossary = data.glossary;
        if (data.figures) figures = data.figures;
        if (data.ganttTasks) ganttTasks = data.ganttTasks;

        // Restore logos
        if (data.logos) {
            logos = data.logos;
            if (logos.logo_ecole) {
                document.getElementById('logoEcolePreview').innerHTML = `<img src="${logos.logo_ecole}" alt="Logo">`;
            }
            if (logos.logo_entreprise) {
                document.getElementById('logoEntreprisePreview').innerHTML = `<img src="${logos.logo_entreprise}" alt="Logo">`;
            }
            if (logos.image_centrale) {
                document.getElementById('imageCentralePreview').innerHTML = `<img src="${logos.image_centrale}" alt="Image">`;
            }
        }

        // Render lists
        renderGlossary();
        renderFigures();
        renderGantt();

    } catch (e) {
        console.error('Error loading from localStorage:', e);
    }
}

function initAutoSave() {
    document.querySelectorAll('input, select, textarea').forEach(el => {
        el.addEventListener('change', saveToLocalStorage);
        el.addEventListener('input', debounce(saveToLocalStorage, 500));
    });
}

function debounce(func, wait) {
    let timeout;
    return function(...args) {
        clearTimeout(timeout);
        timeout = setTimeout(() => func.apply(this, args), wait);
    };
}

function showSaveIndicator() {
    const indicator = document.getElementById('saveIndicator');
    indicator.classList.add('show');
    setTimeout(() => indicator.classList.remove('show'), 1500);
}

// ===== COLLECT FORM DATA =====
function collectFormData() {
    const getValue = (id) => document.getElementById(id)?.value || '';
    const getChecked = (id) => document.getElementById(id)?.checked || false;
    const getRadio = (name) => document.querySelector(`input[name="${name}"]:checked`)?.value || '';

    return {
        cover_model: getRadio('cover_model') || 'classique',
        prenom: getValue('prenom'),
        nom: getValue('nom'),
        formation: getValue('formation'),
        ecole: getValue('ecole'),
        annee_scolaire: getValue('annee_scolaire'),

        entreprise_nom: getValue('entreprise_nom'),
        entreprise_secteur: getValue('entreprise_secteur'),
        entreprise_ville: getValue('entreprise_ville'),
        tuteur_nom: getValue('tuteur_nom'),
        tuteur_poste: getValue('tuteur_poste'),
        tuteur_academique_nom: getValue('tuteur_academique_nom'),
        tuteur_academique_poste: getValue('tuteur_academique_poste'),

        date_debut: getValue('date_debut'),
        date_fin: getValue('date_fin'),
        poste: getValue('poste'),
        mission_principale: getValue('mission_principale'),

        chapters: chapters,
        glossary: glossary,
        figures: figures,
        ganttTasks: ganttTasks,

        include_cover: getChecked('include_cover'),
        include_thanks: getChecked('include_thanks'),
        include_toc: getChecked('include_toc'),
        include_figures_list: getChecked('include_figures_list'),
        include_abstract: getChecked('include_abstract'),
        include_glossary: getChecked('include_glossary'),
        include_gantt: getChecked('include_gantt'),
        include_annexes: getChecked('include_annexes'),

        style: {
            font_family: getValue('font_family'),
            font_size: parseInt(getValue('font_size')) || 12,
            line_spacing: parseFloat(getValue('line_spacing')) || 1.5,
            title1_size: parseInt(getValue('title1_size')) || 16,
            title1_bold: getChecked('title1_bold'),
            title1_color: getValue('title1_color'),
            title2_size: parseInt(getValue('title2_size')) || 14,
            title2_bold: getChecked('title2_bold'),
            title2_color: getValue('title2_color'),
            title3_size: parseInt(getValue('title3_size')) || 12,
            title3_italic: getChecked('title3_italic'),
            title3_color: getValue('title3_color')
        },

        page: {
            margin_top: parseFloat(getValue('margin_top')) || 2.5,
            margin_bottom: parseFloat(getValue('margin_bottom')) || 2.5,
            margin_left: parseFloat(getValue('margin_left')) || 2.5,
            margin_right: parseFloat(getValue('margin_right')) || 2.5,
            show_page_number: getChecked('show_page_number'),
            show_student_name: getChecked('show_student_name')
        },

        logos: logos
    };
}

// ===== LIVE PREVIEW =====
function updatePreview() {
    const data = collectFormData();
    const preview = document.getElementById('documentPreview');
    const primaryColor = data.style.title1_color || '#1a365d';

    let html = '';

    // Page de garde
    if (data.include_cover) {
        const coverModel = data.cover_model || 'classique';

        if (coverModel === 'moderne') {
            // Style Moderne
            html += `<div class="preview-page preview-cover-moderne">`;

            // Bandeau image
            if (data.logos.image_centrale) {
                html += `<div class="preview-banner"><img src="${data.logos.image_centrale}"></div>`;
            } else {
                html += `<div class="preview-banner" style="background: linear-gradient(135deg, ${primaryColor}20, ${primaryColor}40);"></div>`;
            }

            html += `<div class="preview-cover" style="text-align:center;">
                <div style="font-size:18px;font-weight:bold;color:${primaryColor};">RAPPORT</div>
                <div style="font-size:14px;color:#666;">DE STAGE</div>
                ${data.poste ? `<p style="font-size:10px;color:${primaryColor};margin-top:8px;font-style:italic;">${data.poste}</p>` : ''}
                <p style="margin-top:15px;font-size:12px;"><strong>${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</strong></p>
                <p style="font-size:9px;">${data.formation || '[Formation]'}</p>
                <p style="font-size:8px;color:#666;">${data.ecole || '[École]'} • ${data.annee_scolaire || '[Année]'}</p>
            </div>`;

            // Logos côte à côte
            html += `<div style="display:flex;justify-content:center;align-items:center;gap:15px;margin:10px 0;">
                <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                <span style="color:#ccc;">→</span>
                <div class="preview-logo-sm">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
            </div>`;

            html += `<div style="text-align:center;font-size:9px;">
                <p><strong>${data.entreprise_nom || '[Entreprise]'}</strong></p>
                <p style="color:#666;">${data.entreprise_secteur || ''} ${data.entreprise_ville ? '• ' + data.entreprise_ville : ''}</p>
                <p style="color:${primaryColor};margin-top:5px;">${data.date_debut || '[Date]'} — ${data.date_fin || '[Date]'}</p>
            </div>`;

            html += `</div>`;

        } else if (coverModel === 'corporate') {
            // Style Corporate (sidebar)
            html += `<div class="preview-page preview-cover-corporate" style="display:flex;padding:0;">`;

            // Sidebar
            html += `<div class="preview-sidebar" style="width:30%;background:${primaryColor};padding:10px;display:flex;flex-direction:column;align-items:center;justify-content:center;">
                <div class="preview-logo-sm" style="margin-bottom:10px;">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}" style="max-height:25px;">` : ''}</div>
                <div style="color:white;font-size:10px;font-weight:bold;">${data.annee_scolaire || '[Année]'}</div>
                <div class="preview-logo-sm" style="margin-top:10px;">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}" style="max-height:25px;">` : ''}</div>
            </div>`;

            // Contenu principal
            html += `<div style="flex:1;padding:15px 10px;">
                <div style="font-size:16px;font-weight:bold;color:${primaryColor};">RAPPORT</div>
                <div style="font-size:13px;color:#555;">DE STAGE</div>
                <p style="font-size:9px;font-style:italic;margin-top:5px;">${data.formation || '[Formation]'}</p>
                <div style="border-top:2px solid ${primaryColor};width:50%;margin:8px 0;"></div>
                ${data.logos.image_centrale ? `<div style="margin:8px 0;"><img src="${data.logos.image_centrale}" style="max-height:40px;max-width:100%;"></div>` : ''}
                <p style="font-size:7px;color:#888;margin-top:8px;">RÉALISÉ PAR</p>
                <p style="font-size:10px;font-weight:bold;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>
                <p style="font-size:8px;">${data.ecole || '[École]'}</p>
                <p style="font-size:7px;color:#888;margin-top:8px;">ENTREPRISE</p>
                <p style="font-size:9px;font-weight:bold;color:${primaryColor};">${data.entreprise_nom || '[Entreprise]'}</p>
                <p style="font-size:7px;">${data.entreprise_ville || ''}</p>
                <p style="font-size:7px;color:#888;margin-top:8px;">PÉRIODE</p>
                <p style="font-size:8px;">${data.date_debut || '[Date]'} → ${data.date_fin || '[Date]'}</p>
            </div>`;

            html += `</div>`;

        } else {
            // Style Classique (défaut)
            html += `<div class="preview-page">`;

            // Header avec logos
            html += `<div class="preview-header-bar">
                <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                <div class="preview-logo-sm">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
            </div>`;

            // Contenu page de garde
            html += `<div class="preview-cover">
                <div class="preview-cover-title" style="color:${primaryColor};">RAPPORT DE STAGE</div>
                <p style="font-size:10px;font-style:italic;">${data.formation || '[Formation]'}</p>
                <div style="border-top:2px solid ${primaryColor};width:60%;margin:10px auto;"></div>
                ${data.logos.image_centrale ? `<div class="preview-cover-image"><img src="${data.logos.image_centrale}"></div>` : ''}
                <p style="margin-top:10px;font-size:11px;"><strong>${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</strong></p>
                <p style="font-size:9px;">${data.ecole || '[École]'}</p>
                <p style="font-size:8px;font-style:italic;">Année ${data.annee_scolaire || '[Année]'}</p>
                <div style="border-top:2px solid ${primaryColor};width:60%;margin:10px auto;"></div>
                <p style="font-size:8px;">Stage chez</p>
                <p style="font-size:10px;font-weight:bold;color:${primaryColor};">${data.entreprise_nom || '[Entreprise]'}</p>
                <p style="font-size:8px;">${data.entreprise_ville || '[Ville]'}</p>
                <p style="font-size:8px;margin-top:8px;">${data.date_debut || '[Date]'} au ${data.date_fin || '[Date]'}</p>
            </div>`;

            // Tuteurs
            html += `<div style="display:flex;justify-content:space-around;font-size:7px;margin-top:10px;">
                <div style="text-align:center;">
                    <div style="color:#888;">Tuteur entreprise</div>
                    <div><strong>${data.tuteur_nom || '[Nom]'}</strong></div>
                </div>
                <div style="text-align:center;">
                    <div style="color:#888;">Tuteur académique</div>
                    <div><strong>${data.tuteur_academique_nom || '[Nom]'}</strong></div>
                </div>
            </div>`;

            html += `</div>`;
        }
    }

    // Table des matières
    if (data.include_toc) {
        html += `<div class="preview-page">
            <div class="preview-section-title">Table des matières</div>`;

        if (data.include_thanks) {
            html += `<div class="preview-toc-item">Remerciements</div>`;
        }

        let num = 0;
        data.chapters.forEach(ch => {
            num++;
            html += `<div class="preview-toc-item">${num}. ${ch.title}</div>`;
            ch.children.forEach((sub, j) => {
                html += `<div class="preview-toc-item level-2">${num}.${j+1} ${sub.title}</div>`;
            });
        });

        if (data.include_glossary && data.glossary.length > 0) {
            html += `<div class="preview-toc-item">Glossaire</div>`;
        }
        if (data.include_annexes) {
            html += `<div class="preview-toc-item">Annexes</div>`;
        }

        html += `</div>`;
    }

    // Liste des figures
    if (data.include_figures_list && data.figures.length > 0) {
        html += `<div class="preview-page">
            <div class="preview-section-title">Liste des figures</div>`;
        data.figures.forEach((fig, i) => {
            html += `<div class="preview-toc-item">Figure ${i+1} : ${fig.name} <span>p.${fig.page}</span></div>`;
        });
        html += `</div>`;
    }

    // Gantt
    if (data.include_gantt && data.ganttTasks.length > 0) {
        html += `<div class="preview-page">
            <div class="preview-section-title">Diagramme de Gantt</div>
            <div class="preview-gantt">`;

        // Calculer les dates min/max
        const dates = data.ganttTasks.flatMap(t => [new Date(t.start), new Date(t.end)]);
        const minDate = new Date(Math.min(...dates));
        const maxDate = new Date(Math.max(...dates));
        const totalDays = (maxDate - minDate) / (1000 * 60 * 60 * 24) || 1;

        data.ganttTasks.forEach(task => {
            const start = new Date(task.start);
            const end = new Date(task.end);
            const leftPct = ((start - minDate) / (1000 * 60 * 60 * 24)) / totalDays * 100;
            const widthPct = ((end - start) / (1000 * 60 * 60 * 24)) / totalDays * 100;

            html += `<div class="preview-gantt-row">
                <span class="preview-gantt-label">${task.task}</span>
                <div style="flex:1;background:#e5e7eb;border-radius:2px;position:relative;height:10px;">
                    <div class="preview-gantt-bar" style="position:absolute;left:${leftPct}%;width:${widthPct}%;"></div>
                </div>
            </div>`;
        });

        html += `</div></div>`;
    }

    // Glossaire
    if (data.include_glossary && data.glossary.length > 0) {
        html += `<div class="preview-page">
            <div class="preview-section-title">Glossaire</div>`;
        data.glossary.forEach(item => {
            html += `<p style="font-size:9px;"><strong>${item.term}</strong> : ${item.definition}</p>`;
        });
        html += `</div>`;
    }

    preview.innerHTML = html || '<p style="color:#9ca3af;text-align:center;padding:2rem;">Remplissez le formulaire pour voir l\'aperçu</p>';
}

// ===== GENERATE REPORT =====
async function generateReport(format) {
    const btn = document.querySelector(format === 'docx' ? '.btn-primary' : '.btn-pdf');
    const originalText = btn.textContent;

    btn.disabled = true;
    btn.innerHTML = '⏳...';

    try {
        const data = collectFormData();

        const response = await fetch(`/generate?format=${format}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });

        if (!response.ok) throw new Error('Erreur lors de la génération');

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `rapport_stage_${data.nom || 'rapport'}.${format}`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        a.remove();

    } catch (error) {
        alert('Erreur : ' + error.message);
    } finally {
        btn.disabled = false;
        btn.textContent = originalText;
    }
}
