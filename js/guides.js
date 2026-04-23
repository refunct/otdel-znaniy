// guides.js
import { state, elements, escapeHtml, parseMediaList, getFileName, formatExcelDate, navigateTo } from './app.js';

function processContent(content) {
    if (!content) return '';
    
    // 1. Преобразуем списки (уже реализовано)
    let html = convertLists(content);
    
    // 2. Оборачиваем обычный текст в абзацы
    html = wrapParagraphs(html);
    
    return html;
}

// В начало файла после импортов добавим функцию проверки origin
function isSameOrigin(url) {
    try {
        // Относительные пути считаем локальными
        if (url.startsWith('/') || url.startsWith('./') || url.startsWith('../')) return true;
        const parsed = new URL(url, window.location.origin);
        return parsed.origin === window.location.origin;
    } catch {
        return false; // если невалидный URL, считаем внешним
    }
}

// Функция для определения ID видео YouTube
function getYouTubeId(url) {
    const regExp = /^.*(youtu.be\/|v\/|u\/\w\/|embed\/|watch\?v=|&v=)([^#&?]*).*/;
    const match = url.match(regExp);
    return (match && match[2].length === 11) ? match[2] : null;
}

// Вспомогательная функция для оборачивания текста в параграфы
function wrapParagraphs(html) {
    if (!html) return '';
    
    // Сначала защитим блочные теги, чтобы они не ломались при разбиении
    const blockTags = ['ul', 'ol', 'q', 'c', 'w', 'e', 'div', 'blockquote', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
    const placeholder = '%%BLOCK_TAG%%';
    const blocks = [];
    
    // Временно заменяем блочные теги на плейсхолдеры
    let index = 0;
    let processed = html.replace(/<(ul|ol|q|c|w|e|div|blockquote|h[1-6])[\s\S]*?<\/\1>/gi, (match) => {
        blocks.push(match);
        return placeholder + (index++);
    });
    
    // Разделяем по двум и более пустым строкам (это будущие абзацы)
    const paragraphs = processed.split(/\n\s*\n\s*\n+/);
    const result = [];
    
    for (let para of paragraphs) {
        para = para.trim();
        if (!para) continue;
        
        // Восстанавливаем блочные теги из плейсхолдеров
        para = para.replace(new RegExp(placeholder + '(\\d+)', 'g'), (_, i) => blocks[parseInt(i)]);
        
        // Если абзац начинается с блочного тега — не оборачиваем в <p>
        const isBlockTag = /^\s*<(ul|ol|q|c|w|e|div|blockquote|h[1-6])/i.test(para);
        if (isBlockTag) {
            result.push(para);
        } else {
            // Внутри абзаца:
            // 1. Два переноса подряд (одна пустая строка) -> <br>
            // 2. Одиночный перенос -> пробел
            let content = para
                .replace(/\n\n/g, '<br>')
                .replace(/\n/g, ' ');
            
            result.push(`<p>${content}</p>`);
        }
    }
    
    return result.join('\n');
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
                        ${images.map(img => `<img src="${escapeHtml(img)}" class="section-image" loading="lazy" data-lightbox onerror="if(this.parentElement) this.parentElement.innerHTML='<div class=\\'offline-placeholder\\'>Не удалось загрузить</div>'">`).join('')}
                    </div>
                </div>` : ''}
            ${files.length ? `
                <div class="media-section"><div class="media-title">Файлы</div>
                    <ul class="files-list">
                        ${files.map(f => `<li class="file-item"><a href="${escapeHtml(f)}" class="file-link" download>📄 ${escapeHtml(getFileName(f))}</a></li>`).join('')}
                    </ul>
                </div>` : ''}
            ${videos.length ? `
                <div class="media-section">
                    <div class="media-title">Видео</div>
                    <div class="videos-grid">
                        ${videos.map(video => {
                            const sameOrigin = isSameOrigin(video);
                            const youtubeId = !sameOrigin ? getYouTubeId(video) : null;
                            if (sameOrigin) {
                                return `<video controls class="video-player" src="${escapeHtml(video)}"></video>`;
                            } else if (youtubeId) {
                                return `<iframe class="video-iframe" src="https://www.youtube.com/embed/${youtubeId}" frameborder="0" allowfullscreen></iframe>`;
                            } else {
                                return `<iframe class="video-iframe" src="${escapeHtml(video)}" frameborder="0" allowfullscreen></iframe>`;
                            }
                        }).join('')}
                    </div>
                </div>
            ` : ''}
        </div>
    `;
}

function convertLists(content) {
    if (!content) return '';
    
    const lines = content.split('\n');
    const result = [];
    let listItems = [];
    let listType = null; // 'ul' или 'ol'
    
    const flushList = () => {
        if (listItems.length > 0) {
            const tag = listType === 'ul' ? 'ul' : 'ol';
            result.push(`<${tag}>${listItems.join('')}</${tag}>`);
            listItems = [];
            listType = null;
        }
    };
    
    for (let line of lines) {
        // Маркированный список: начинается с "- " или "* "
        const ulMatch = line.match(/^(\s*)([-*])\s+(.*)$/);
        // Нумерованный список: начинается с цифр и точки "1. "
        const olMatch = line.match(/^(\s*)(\d+)\.\s+(.*)$/);
        
        if (ulMatch) {
            const indent = ulMatch[1].length;
            const text = ulMatch[3];
            if (listType !== 'ul') {
                flushList();
                listType = 'ul';
            }
            // TODO: можно добавить поддержку вложенности по отступам, пока просто плоский список
            listItems.push(`<li>${escapeHtml(text)}</li>`);
        } else if (olMatch) {
            const indent = olMatch[1].length;
            const text = olMatch[3];
            if (listType !== 'ol') {
                flushList();
                listType = 'ol';
            }
            listItems.push(`<li>${escapeHtml(text)}</li>`);
        } else {
            // Не список – сбрасываем накопленный список
            flushList();
            result.push(line);
        }
    }
    flushList();
    
    return result.join('\n');
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