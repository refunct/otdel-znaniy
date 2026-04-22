// guides.js
import { state, elements, escapeHtml, parseMediaList, getFileName, formatExcelDate, navigateTo } from './app.js';

function processContent(content) {
    if (!content) return '';
    // Заменяем кастомные теги на span с классами для стилизации
    return content
        .replace(/<c>/g, '<c>').replace(/<\/c>/g, '</c>')
        .replace(/<w>/g, '<w>').replace(/<\/w>/g, '</w>')
        .replace(/<e>/g, '<e>').replace(/<\/e>/g, '</e>');
}

function renderSection(section) {
    const images = parseMediaList(section.images);
    const files = parseMediaList(section.files);
    const videos = parseMediaList(section.videos);
    return `
        <div class="section-card">
            <h3 class="section-title">${escapeHtml(section.title || 'Без названия')}</h3>
            <div class="section-content">${processContent(section.content || '')}</div>
            ${images.length ? `
                <div class="media-section"><div class="media-title">Изображения</div>
                    <div class="images-grid">
                        ${images.map(img => `<img src="${escapeHtml(img)}" class="section-image" loading="lazy" onerror="this.parentElement.innerHTML='<div class=\\'offline-placeholder\\'>Не удалось загрузить</div>'">`).join('')}
                    </div>
                </div>` : ''}
            ${files.length ? `
                <div class="media-section"><div class="media-title">Файлы</div>
                    <ul class="files-list">
                        ${files.map(f => `<li class="file-item"><a href="${escapeHtml(f)}" class="file-link" download>📄 ${escapeHtml(getFileName(f))}</a></li>`).join('')}
                    </ul>
                </div>` : ''}
            ${videos.length ? `
                <div class="media-section"><div class="media-title">Видео</div>
                    <ul class="videos-list">
                        ${videos.map(v => `<li class="video-item"><a href="${escapeHtml(v)}" class="video-link" target="_blank">🎬 ${escapeHtml(getFileName(v))}</a></li>`).join('')}
                    </ul>
                </div>` : ''}
        </div>
    `;
}

export function renderGuidesList() {
    if (!state.guides.length) {
        elements.contentContainer.innerHTML = '<div class="empty-state">Справочники не найдены</div>';
        return;
    }
    const html = `<div class="guides-grid">${state.guides.map(g => `
        <div class="guide-card" data-guide-id="${g.id}">
            ${g.image ? `<img src="${escapeHtml(g.image)}" class="card-image" loading="lazy">` : '<div class="card-image"></div>'}
            <div class="card-content">
                <h3 class="card-title">${escapeHtml(g.title)}</h3>
                <div class="card-meta">${g.author ? `<span>${escapeHtml(g.author)}</span>` : ''} ${g.date ? `<span>${formatExcelDate(g.date)}</span>` : ''}</div>
            </div>
        </div>`).join('')}</div>`;
    elements.contentContainer.innerHTML = html;
    document.querySelectorAll('.guide-card').forEach(c => c.addEventListener('click', () => navigateTo('guides', c.dataset.guideId)));
}

export function renderSections() {
    const guide = state.guides.find(g => g.id === state.currentGuideId);
    if (!guide) return navigateTo('guides');
    const sections = state.sections.filter(s => s.guide_id === state.currentGuideId);
    const html = `
        <div class="sections-container">
            <button class="back-button" id="backToGuides">← Назад</button>
            <h2>${escapeHtml(guide.title)}</h2>
            ${sections.length ? sections.map(renderSection).join('') : '<div class="empty-state">Нет разделов</div>'}
        </div>`;
    elements.contentContainer.innerHTML = html;
    document.getElementById('backToGuides').addEventListener('click', () => navigateTo('guides'));
}