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

            // Niveau 3 : sous-sous-chapitres
            if (sub.children) {
                sub.children.forEach((subsub, subsubIndex) => {
                    container.appendChild(createChapterElement(subsub, subsubIndex, `${chapterNum}.${subIndex + 1}.${subsubIndex + 1}`, index, subIndex));
                });
            }
        });
    });

    saveToLocalStorage();
    updatePreview();
}

function createChapterElement(chapter, index, number, parentIndex = null, grandParentSubIndex = null) {
    const div = document.createElement('div');
    div.className = `chapter-item level-${chapter.level}`;

    // Permettre d'ajouter des sous-éléments jusqu'au niveau 2 (pour créer niveau 3)
    const canAddSub = chapter.level < 3;
    const canMove = parentIndex === null;

    div.innerHTML = `
        <span class="chapter-number">${number}</span>
        <input type="text" class="chapter-input" value="${chapter.title}"
               onchange="updateChapterTitle(${chapter.id}, this.value)">
        <div class="chapter-actions">
            ${canAddSub ? `<button class="add-sub" onclick="addSubChapter(${chapter.id}, ${chapter.level})">+</button>` : ''}
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

function addSubChapter(parentId, parentLevel) {
    const newLevel = parentLevel + 1;
    const newTitle = newLevel === 2 ? 'Nouvelle section' : 'Nouvelle sous-section';

    function findAndAdd(items) {
        for (let item of items) {
            if (item.id === parentId) {
                if (!item.children) item.children = [];
                item.children.push({ id: Date.now(), title: newTitle, level: newLevel, children: [] });
                return true;
            }
            if (item.children && findAndAdd(item.children)) return true;
        }
        return false;
    }

    findAndAdd(chapters);
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
            'date_debut', 'date_fin', 'sujet_stage', 'poste',
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
        sujet_stage: getValue('sujet_stage'),
        poste: getValue('poste'),

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
                ${data.sujet_stage ? `<p style="font-size:11px;font-weight:bold;color:${primaryColor};margin-top:8px;">${data.sujet_stage}</p>` : ''}
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

        } else if (coverModel === 'elegant') {
            // Style Elegant (ligne verticale)
            html += `<div class="preview-page" style="display:flex;padding:0;">`;

            // Ligne verticale
            html += `<div style="width:8px;background:${primaryColor};"></div>`;

            // Contenu principal
            html += `<div style="flex:1;padding:15px 12px;">
                <div style="display:flex;justify-content:space-between;margin-bottom:15px;">
                    <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                    <div class="preview-logo-sm">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
                </div>
                <div style="font-size:18px;font-weight:bold;color:${primaryColor};">RAPPORT</div>
                <div style="font-size:12px;color:#666;margin-bottom:8px;">DE STAGE</div>
                ${data.sujet_stage ? `<p style="font-size:9px;font-weight:bold;font-style:italic;margin-bottom:10px;">${data.sujet_stage}</p>` : ''}
                <div style="border-top:1px solid ${primaryColor};width:80%;margin:10px 0;"></div>
                <p style="font-size:11px;font-weight:bold;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>
                <p style="font-size:8px;font-style:italic;">${data.formation || '[Formation]'}</p>
                <p style="font-size:7px;color:#888;margin-top:3px;">${data.ecole || '[École]'}  |  ${data.annee_scolaire || '[Année]'}</p>
                <p style="font-size:9px;font-weight:bold;color:${primaryColor};margin-top:12px;">${data.entreprise_nom || '[Entreprise]'}</p>
                <p style="font-size:7px;color:#666;">${data.entreprise_ville || ''}</p>
                <p style="font-size:8px;font-weight:bold;margin-top:10px;">${data.date_debut || '[Date]'}  →  ${data.date_fin || '[Date]'}</p>
            </div>`;

            html += `</div>`;

        } else if (coverModel === 'minimaliste') {
            // Style Minimaliste - Ultra epure
            html += `<div class="preview-page" style="display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;">`;
            html += `<div style="margin-bottom:30px;">
                <div style="font-size:18px;font-weight:bold;color:${primaryColor};">RAPPORT DE STAGE</div>
                <div style="font-size:10px;color:${primaryColor};margin-top:5px;">─────────────</div>
            </div>`;
            if (data.sujet_stage) {
                html += `<p style="font-size:11px;font-style:italic;margin-bottom:20px;">${data.sujet_stage}</p>`;
            }
            html += `<p style="font-size:12px;font-weight:bold;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>`;
            html += `<p style="font-size:9px;color:#666;">${data.formation || '[Formation]'}</p>`;
            html += `<div style="margin-top:40px;font-size:8px;color:#888;">
                <p>${data.entreprise_nom || '[Entreprise]'}</p>
                <p>${data.date_debut || '[Date]'} — ${data.date_fin || '[Date]'}</p>
            </div>`;
            html += `</div>`;

        } else if (coverModel === 'academique') {
            // Style Academique - Cadre double
            html += `<div class="preview-page" style="padding:8px;">`;
            html += `<div style="border:3px solid ${primaryColor};padding:6px;">`;
            html += `<div style="border:1px solid ${primaryColor};padding:12px;text-align:center;">`;

            // Logos
            html += `<div style="display:flex;justify-content:space-between;margin-bottom:10px;">
                <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                <div class="preview-logo-sm">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
            </div>`;

            html += `<p style="font-size:10px;font-weight:bold;color:${primaryColor};">${data.ecole || '[École]'}</p>`;
            html += `<p style="font-size:8px;margin-bottom:10px;">${data.formation || '[Formation]'}</p>`;
            html += `<div style="font-size:16px;font-weight:bold;color:${primaryColor};margin:15px 0;">RAPPORT DE STAGE</div>`;
            if (data.sujet_stage) {
                html += `<p style="font-size:10px;font-style:italic;margin-bottom:10px;">${data.sujet_stage}</p>`;
            }
            html += `<p style="font-size:8px;color:#666;margin:10px 0;">Présenté par</p>`;
            html += `<p style="font-size:11px;font-weight:bold;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>`;
            html += `<p style="font-size:8px;margin-top:10px;">Stage chez ${data.entreprise_nom || '[Entreprise]'}</p>`;
            html += `<p style="font-size:7px;color:#666;">${data.date_debut || '[Date]'} au ${data.date_fin || '[Date]'}</p>`;
            html += `<p style="font-size:9px;color:${primaryColor};margin-top:15px;">${data.annee_scolaire || '[Année]'}</p>`;

            html += `</div></div></div>`;

        } else if (coverModel === 'geometrique') {
            // Style Geometrique - Formes modernes
            html += `<div class="preview-page" style="position:relative;overflow:hidden;">`;

            // Forme geometrique en haut a droite
            html += `<div style="position:absolute;top:-20px;right:-20px;width:80px;height:80px;background:linear-gradient(135deg,${primaryColor} 50%,transparent 50%);"></div>`;

            // Bloc colore en haut
            html += `<div style="display:flex;margin-bottom:20px;">
                <div style="flex:1;">
                    <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                </div>
                <div style="background:${primaryColor};padding:8px 15px;color:white;font-size:9px;font-weight:bold;">
                    ${data.annee_scolaire || '[Année]'}
                </div>
            </div>`;

            html += `<div style="font-size:18px;font-weight:bold;color:${primaryColor};">RAPPORT DE STAGE</div>`;
            if (data.sujet_stage) {
                html += `<p style="font-size:10px;font-style:italic;margin:8px 0;">${data.sujet_stage}</p>`;
            }
            html += `<div style="width:60px;height:4px;background:${primaryColor};margin:15px 0;"></div>`;

            html += `<p style="font-size:11px;font-weight:bold;margin-top:20px;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>`;
            html += `<p style="font-size:8px;">${data.formation || '[Formation]'}</p>`;
            html += `<p style="font-size:7px;color:#888;">${data.ecole || '[École]'}</p>`;

            html += `<div style="margin-top:25px;font-size:8px;">
                <p><span style="color:#888;">Entreprise :</span> <strong>${data.entreprise_nom || '[Entreprise]'}</strong></p>
                <p style="color:#666;">${data.date_debut || '[Date]'} au ${data.date_fin || '[Date]'}</p>
            </div>`;

            html += `</div>`;

        } else if (coverModel === 'bicolore') {
            // Style Bicolore - Split vertical
            html += `<div class="preview-page" style="display:flex;padding:0;">`;

            // Colonne gauche coloree
            html += `<div style="width:35%;background:${primaryColor};padding:12px;color:white;display:flex;flex-direction:column;align-items:center;justify-content:center;">`;
            if (data.logos.logo_ecole) {
                html += `<div class="preview-logo-sm" style="background:white;margin-bottom:15px;"><img src="${data.logos.logo_ecole}"></div>`;
            }
            html += `<div style="font-size:14px;font-weight:bold;margin:20px 0;">STAGE</div>`;
            html += `<p style="font-size:8px;opacity:0.8;">${data.annee_scolaire || '[Année]'}</p>`;
            html += `<p style="font-size:7px;opacity:0.7;margin-top:10px;text-align:center;">${data.date_debut || '[Date]'}<br>—<br>${data.date_fin || '[Date]'}</p>`;
            if (data.logos.logo_entreprise) {
                html += `<div class="preview-logo-sm" style="background:white;margin-top:20px;"><img src="${data.logos.logo_entreprise}"></div>`;
            }
            html += `</div>`;

            // Colonne droite
            html += `<div style="flex:1;padding:15px;display:flex;flex-direction:column;justify-content:center;">`;
            html += `<div style="font-size:16px;font-weight:bold;color:${primaryColor};line-height:1.2;">RAPPORT<br>DE STAGE</div>`;
            if (data.sujet_stage) {
                html += `<p style="font-size:9px;font-style:italic;margin:10px 0;">${data.sujet_stage}</p>`;
            }
            html += `<p style="font-size:11px;font-weight:bold;margin-top:20px;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>`;
            html += `<p style="font-size:8px;">${data.formation || '[Formation]'}</p>`;
            html += `<p style="font-size:7px;color:#888;">${data.ecole || '[École]'}</p>`;
            html += `<p style="font-size:10px;font-weight:bold;color:${primaryColor};margin-top:15px;">${data.entreprise_nom || '[Entreprise]'}</p>`;
            html += `</div>`;

            html += `</div>`;

        } else if (coverModel === 'pro') {
            // Style Pro - Corporate business
            html += `<div class="preview-page" style="padding:0;">`;

            // Header bar
            html += `<div style="background:${primaryColor};padding:10px;display:flex;justify-content:space-between;align-items:center;">
                <div class="preview-logo-sm" style="background:white;">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                <span style="color:white;font-size:8px;">${data.annee_scolaire || '[Année]'}</span>
                <div class="preview-logo-sm" style="background:white;">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
            </div>`;

            // Contenu central
            html += `<div style="padding:20px;text-align:center;">`;
            html += `<div style="font-size:18px;font-weight:bold;color:${primaryColor};">RAPPORT DE STAGE</div>`;
            if (data.sujet_stage) {
                html += `<p style="font-size:10px;font-style:italic;margin:10px 0;">${data.sujet_stage}</p>`;
            }
            html += `<p style="font-size:11px;font-weight:bold;margin-top:25px;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>`;
            html += `<p style="font-size:8px;color:#666;">${data.formation || '[Formation]'}  •  ${data.ecole || '[École]'}</p>`;

            // Info table
            html += `<div style="margin-top:20px;text-align:left;font-size:8px;">
                <div style="display:flex;margin:5px 0;"><span style="color:#888;width:60px;">Entreprise</span><strong>${data.entreprise_nom || '[Entreprise]'}</strong></div>
                <div style="display:flex;margin:5px 0;"><span style="color:#888;width:60px;">Période</span>${data.date_debut || '[Date]'} — ${data.date_fin || '[Date]'}</div>
                <div style="display:flex;margin:5px 0;"><span style="color:#888;width:60px;">Tuteur</span>${data.tuteur_nom || '[Nom]'}</div>
            </div>`;
            html += `</div>`;

            // Footer bar
            html += `<div style="border-top:3px solid ${primaryColor};padding:8px;text-align:center;">
                <span style="font-size:7px;color:#888;">${data.entreprise_nom || ''} ${data.entreprise_ville ? '— ' + data.entreprise_ville : ''}</span>
            </div>`;

            html += `</div>`;

        } else if (coverModel === 'gradient') {
            // Style Gradient - Degrade colore
            html += `<div class="preview-page" style="padding:0;">`;

            // Bandeau degrade
            html += `<div style="background:linear-gradient(135deg, ${primaryColor} 0%, #667eea 50%, #764ba2 100%);padding:20px;text-align:center;">`;
            html += `<div style="font-size:16px;font-weight:bold;color:white;">RAPPORT DE STAGE</div>`;
            if (data.sujet_stage) {
                html += `<p style="font-size:9px;font-style:italic;color:rgba(255,255,255,0.9);margin-top:8px;">${data.sujet_stage}</p>`;
            }
            html += `</div>`;

            // Logos
            html += `<div style="display:flex;justify-content:center;gap:20px;margin:15px 0;">
                <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                <div class="preview-logo-sm">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
            </div>`;

            // Contenu
            html += `<div style="padding:15px;text-align:center;">`;
            html += `<p style="font-size:11px;font-weight:bold;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>`;
            html += `<p style="font-size:8px;color:#666;">${data.formation || '[Formation]'}</p>`;
            html += `<p style="font-size:7px;color:#888;">${data.ecole || '[École]'}  •  ${data.annee_scolaire || '[Année]'}</p>`;
            html += `<p style="font-size:9px;font-weight:bold;color:${primaryColor};margin-top:15px;">${data.entreprise_nom || '[Entreprise]'}</p>`;
            html += `<p style="font-size:7px;color:#666;">${data.date_debut || '[Date]'} — ${data.date_fin || '[Date]'}</p>`;
            html += `</div>`;

            html += `</div>`;

        } else if (coverModel === 'timeline') {
            // Style Timeline - Frise temporelle
            html += `<div class="preview-page" style="display:flex;padding:0;">`;

            // Colonne timeline
            html += `<div style="width:25px;padding:15px 5px;display:flex;flex-direction:column;align-items:center;">
                <div style="width:8px;height:8px;background:${primaryColor};border-radius:50%;margin-bottom:5px;"></div>
                <div style="font-size:5px;color:#888;margin-bottom:8px;">${data.date_debut ? data.date_debut.split('-')[2] : ''}</div>
                <div style="flex:1;width:2px;background:linear-gradient(to bottom,${primaryColor},#818cf8);"></div>
                <div style="font-size:5px;color:#888;margin-top:8px;">${data.date_fin ? data.date_fin.split('-')[2] : ''}</div>
                <div style="width:8px;height:8px;background:#818cf8;border-radius:50%;margin-top:5px;"></div>
            </div>`;

            // Contenu
            html += `<div style="flex:1;padding:15px 10px;">`;

            // Logos
            html += `<div style="display:flex;justify-content:space-between;margin-bottom:15px;">
                <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                <div class="preview-logo-sm">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
            </div>`;

            html += `<div style="font-size:16px;font-weight:bold;color:${primaryColor};">RAPPORT DE STAGE</div>`;
            if (data.sujet_stage) {
                html += `<p style="font-size:9px;font-style:italic;margin:8px 0;">${data.sujet_stage}</p>`;
            }
            html += `<p style="font-size:10px;font-weight:bold;margin-top:20px;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>`;
            html += `<p style="font-size:7px;color:#888;">${data.formation || '[Formation]'}</p>`;
            html += `<p style="font-size:9px;font-weight:bold;color:${primaryColor};margin-top:15px;">${data.entreprise_nom || '[Entreprise]'}</p>`;
            html += `</div>`;

            html += `</div>`;

        } else if (coverModel === 'creative') {
            // Style Creative - Design original
            html += `<div class="preview-page" style="position:relative;overflow:hidden;">`;

            // Cercles decoratifs
            html += `<div style="position:absolute;top:-20px;right:-20px;width:60px;height:60px;background:linear-gradient(135deg,#f093fb,#f5576c);border-radius:50%;opacity:0.7;"></div>`;
            html += `<div style="position:absolute;bottom:-15px;left:-15px;width:40px;height:40px;background:linear-gradient(135deg,#4facfe,#00f2fe);border-radius:50%;opacity:0.5;"></div>`;

            // Logos
            html += `<div style="display:flex;justify-content:space-between;margin-bottom:20px;position:relative;z-index:1;">
                <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                <div class="preview-logo-sm">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
            </div>`;

            // Titre stylise
            html += `<div style="text-align:center;position:relative;z-index:1;">`;
            html += `<div style="font-size:20px;font-weight:bold;color:${primaryColor};">RAPPORT</div>`;
            html += `<div style="font-size:12px;color:#999;">DE STAGE</div>`;
            html += `<div style="width:50%;height:2px;background:linear-gradient(90deg,#f093fb,#f5576c);margin:10px auto;"></div>`;

            if (data.sujet_stage) {
                html += `<p style="font-size:9px;font-style:italic;margin:10px 0;">« ${data.sujet_stage} »</p>`;
            }

            html += `<p style="font-size:12px;font-weight:bold;margin-top:25px;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>`;
            html += `<p style="font-size:8px;color:#888;">${data.formation || '[Formation]'}</p>`;
            html += `<p style="font-size:7px;">${data.ecole || '[École]'}  ✦  ${data.annee_scolaire || '[Année]'}</p>`;

            html += `<div style="width:30%;height:1px;background:#ddd;margin:15px auto;"></div>`;
            html += `<p style="font-size:9px;font-weight:bold;color:${primaryColor};">${data.entreprise_nom || '[Entreprise]'}</p>`;
            html += `<p style="font-size:7px;color:#888;">${data.date_debut || '[Date]'}  →  ${data.date_fin || '[Date]'}</p>`;
            html += `</div>`;

            html += `</div>`;

        } else if (coverModel === 'luxe') {
            // Style Luxe - Elegant avec bordures dorees
            html += `<div class="preview-page" style="padding:6px;background:#fdfbfb;">`;

            // Double bordure doree
            html += `<div style="border:2px solid #b8860b;padding:4px;">`;
            html += `<div style="border:1px solid #b8860b;padding:12px;">`;

            // Logos
            html += `<div style="display:flex;justify-content:space-between;margin-bottom:15px;">
                <div class="preview-logo-sm">${data.logos.logo_ecole ? `<img src="${data.logos.logo_ecole}">` : ''}</div>
                <div class="preview-logo-sm">${data.logos.logo_entreprise ? `<img src="${data.logos.logo_entreprise}">` : ''}</div>
            </div>`;

            // Ligne decorative
            html += `<div style="display:flex;align-items:center;justify-content:center;margin:10px 0;">
                <div style="width:30px;height:1px;background:#b8860b;"></div>
                <div style="width:6px;height:6px;border:1px solid #b8860b;transform:rotate(45deg);margin:0 8px;"></div>
                <div style="width:30px;height:1px;background:#b8860b;"></div>
            </div>`;

            // Titre
            html += `<div style="text-align:center;">
                <div style="font-size:18px;font-weight:bold;color:#b8860b;letter-spacing:2px;">RAPPORT DE STAGE</div>
            </div>`;

            if (data.sujet_stage) {
                html += `<p style="text-align:center;font-size:9px;font-style:italic;margin:10px 0;">« ${data.sujet_stage} »</p>`;
            }

            // Separateur
            html += `<div style="width:60%;height:1px;background:linear-gradient(90deg,transparent,#b8860b,transparent);margin:12px auto;"></div>`;

            // Infos etudiant
            html += `<div style="text-align:center;">
                <p style="font-size:11px;font-weight:bold;">${data.prenom || '[Prénom]'} ${data.nom || '[Nom]'}</p>
                <p style="font-size:8px;color:#666;">${data.formation || '[Formation]'}</p>
                <p style="font-size:7px;color:#888;">${data.ecole || '[École]'}  |  ${data.annee_scolaire || '[Année]'}</p>
            </div>`;

            // Entreprise
            html += `<div style="text-align:center;margin-top:15px;">
                <p style="font-size:9px;font-weight:bold;color:#b8860b;">${data.entreprise_nom || '[Entreprise]'}</p>
                <p style="font-size:7px;color:#666;">${data.entreprise_ville || ''}</p>
                <p style="font-size:7px;margin-top:5px;">${data.date_debut || '[Date]'} — ${data.date_fin || '[Date]'}</p>
            </div>`;

            html += `</div></div></div>`;

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
                ${data.sujet_stage ? `<p style="font-size:11px;font-weight:bold;color:${primaryColor};margin:8px 0;">${data.sujet_stage}</p>` : ''}
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
                // Niveau 3 : sous-sous-chapitres
                if (sub.children) {
                    sub.children.forEach((subsub, k) => {
                        html += `<div class="preview-toc-item level-3">${num}.${j+1}.${k+1} ${subsub.title}</div>`;
                    });
                }
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
async function generateReport() {
    const btn = document.querySelector('.btn-primary');
    const originalText = btn.textContent;

    btn.disabled = true;
    btn.innerHTML = 'Generation...';

    try {
        const data = collectFormData();

        const response = await fetch('/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });

        if (!response.ok) throw new Error('Erreur lors de la génération');

        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `rapport_stage_${data.nom || 'rapport'}.docx`;
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
