// src/js/main.js

/**
 * 主应用逻辑
 * 处理用户界面交互和转换流程
 */

// 全局变量
let currentConverter = null;
let isConverting = false;

// DOM元素缓存
const elements = {
    htmlInput: null,
    conversionModeSelect: null,
    pptTitleInput: null,
    preserveBackgroundCheckbox: null,
    extractTextCheckbox: null,
    convertBtn: null,
    previewBtn: null,
    loadExampleBtn: null,
    btnSpinner: null,
    notificationArea: null,
    pptxgenStatus: null,
    html2canvasStatus: null
};

/**
 * 初始化应用
 */
document.addEventListener('DOMContentLoaded', () => {
    initializeElements();
    checkLibraries();
    setupEventListeners();
    loadSavedSettings();
    handleModeChange();
});

/**
 * 初始化DOM元素引用
 */
function initializeElements() {
    elements.htmlInput = document.getElementById('htmlInput');
    elements.conversionModeSelect = document.getElementById('conversionModeSelect');
    elements.pptTitleInput = document.getElementById('pptTitleInput');
    elements.preserveBackgroundCheckbox = document.getElementById('preserveBackgroundCheckbox');
    elements.extractTextCheckbox = document.getElementById('extractTextCheckbox');
    elements.convertBtn = document.getElementById('convertBtn');
    elements.previewBtn = document.getElementById('previewBtn');
    elements.loadExampleBtn = document.getElementById('loadExampleBtn');
    elements.btnSpinner = elements.convertBtn ? elements.convertBtn.querySelector('.btn-spinner') : null;
    elements.notificationArea = document.getElementById('notificationArea');
    elements.pptxgenStatus = document.getElementById('pptxgenjs-status');
    elements.html2canvasStatus = document.getElementById('html2canvas-status');
}

/**
 * 检查依赖库加载状态
 */
function checkLibraries() {
    let allLibrariesLoaded = true;

    // 检查 PptxGenJS
    if (typeof PptxGenJS !== 'undefined') {
        elements.pptxgenStatus.textContent = '已加载';
        elements.pptxgenStatus.className = 'library-status loaded';
    } else {
        elements.pptxgenStatus.textContent = '加载失败';
        elements.pptxgenStatus.className = 'library-status error';
        allLibrariesLoaded = false;
        console.error('PptxGenJS未能加载。');
        showNotification('PptxGenJS未能加载，部分功能可能无法使用。', 'error');
    }

    // 检查 html2canvas
    if (typeof html2canvas !== 'undefined') {
        elements.html2canvasStatus.textContent = '已加载';
        elements.html2canvasStatus.className = 'library-status loaded';
    } else {
        elements.html2canvasStatus.textContent = '加载失败';
        elements.html2canvasStatus.className = 'library-status error';
        allLibrariesLoaded = false;
        console.error('html2canvas未能加载。');
        showNotification('html2canvas未能加载，部分功能可能无法使用。', 'error');
    }

    // 根据库加载情况启用/禁用转换按钮
    if (elements.convertBtn) {
        elements.convertBtn.disabled = !allLibrariesLoaded;
        if (!allLibrariesLoaded) {
            showNotification('部分核心库加载失败，转换功能可能无法正常使用。', 'error');
        }
    } else {
        console.error("convertBtn 元素未找到，无法设置其状态。");
    }
}

/**
 * 设置事件监听器
 */
function setupEventListeners() {
    if (elements.conversionModeSelect) {
        elements.conversionModeSelect.addEventListener('change', handleModeChange);
        elements.conversionModeSelect.addEventListener('change', saveSettings);
    }
    if (elements.preserveBackgroundCheckbox) {
        elements.preserveBackgroundCheckbox.addEventListener('change', saveSettings);
    }
    if (elements.extractTextCheckbox) {
        elements.extractTextCheckbox.addEventListener('change', saveSettings);
    }
    if (elements.htmlInput) {
        elements.htmlInput.addEventListener('input', autoResizeTextarea);
    }
    if (elements.convertBtn) {
        elements.convertBtn.addEventListener('click', convertToPPT);
    }
    if (elements.previewBtn) {
        // elements.previewBtn.addEventListener('click', previewHTML); // 假设有 previewHTML 函数
    }
    if (elements.loadExampleBtn) {
        // elements.loadExampleBtn.addEventListener('click', loadSampleHTML); // 假设有 loadSampleHTML 函数
    }
    document.addEventListener('keydown', handleKeyboardShortcuts);
}

/**
 * 处理转换模式改变
 */
function handleModeChange() {
    if (!elements.conversionModeSelect || !elements.preserveBackgroundCheckbox || !elements.extractTextCheckbox) return;
    const mode = elements.conversionModeSelect.value;
    switch(mode) {
        case 'screenshot':
            // 截图模式：默认不提取文字，但允许用户选择
            elements.extractTextCheckbox.disabled = false;
            elements.extractTextCheckbox.checked = false;
            elements.preserveBackgroundCheckbox.checked = true;
            elements.preserveBackgroundCheckbox.disabled = true;
            showNotification('截图模式：勾选"提取可编辑文本"可同时保留背景图片和可编辑文字', 'info');
            break;
        case 'elements':
            elements.preserveBackgroundCheckbox.checked = false;
            elements.preserveBackgroundCheckbox.disabled = true;
            elements.extractTextCheckbox.checked = true;
            elements.extractTextCheckbox.disabled = false;
            showNotification('元素模式：所有内容转换为PPT元素，完全可编辑但可能有样式偏差', 'info');
            break;
        case 'hybrid':
        default:
            // 混合模式：用户可以选择是否提取文字和保留背景
            elements.extractTextCheckbox.disabled = false;
            elements.extractTextCheckbox.checked = true;
            elements.preserveBackgroundCheckbox.disabled = false;
            elements.preserveBackgroundCheckbox.checked = true;
            showNotification('混合模式：可灵活选择背景截图和文字提取的组合（推荐）', 'info');
            break;
    }
}

/**
 * 自动调整文本框高度
 */
function autoResizeTextarea() {
    if (!elements.htmlInput) return;
    elements.htmlInput.style.height = 'auto';
    elements.htmlInput.style.height = Math.min(elements.htmlInput.scrollHeight, 600) + 'px';
}

/**
 * 处理键盘快捷键
 */
function handleKeyboardShortcuts(e) {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        if (elements.convertBtn && !elements.convertBtn.disabled) {
            convertToPPT();
        }
    }
}

/**
 * 保存用户设置
 */
function saveSettings() {
    if (!elements.conversionModeSelect || !elements.preserveBackgroundCheckbox || !elements.extractTextCheckbox || !elements.pptTitleInput) return;
    const settings = {
        mode: elements.conversionModeSelect.value,
        preserveBackground: elements.preserveBackgroundCheckbox.checked,
        extractText: elements.extractTextCheckbox.checked,
        pptTitle: elements.pptTitleInput.value
    };
    localStorage.setItem('html2ppt_settings', JSON.stringify(settings));
}

/**
 * 加载保存的设置
 */
function loadSavedSettings() {
    if (!elements.conversionModeSelect || !elements.preserveBackgroundCheckbox || !elements.extractTextCheckbox || !elements.pptTitleInput) return;
    try {
        const saved = localStorage.getItem('html2ppt_settings');
        if (saved) {
            const settings = JSON.parse(saved);
            elements.conversionModeSelect.value = settings.mode || 'hybrid';
            elements.preserveBackgroundCheckbox.checked = settings.preserveBackground !== false;
            elements.extractTextCheckbox.checked = settings.extractText !== false;
            elements.pptTitleInput.value = settings.pptTitle || '转换的演示文稿';
        }
    } catch (error) {
        console.error('加载设置失败:', error);
        showNotification('加载用户设置失败。', 'error');
    }
}

/**
 * 主转换函数
 */
async function convertToPPT() {
    if (isConverting) {
        showNotification('转换正在进行中，请稍候...', 'warning');
        return;
    }
    
    if (typeof PptxGenJS === 'undefined' || typeof html2canvas === 'undefined') {
        showNotification('核心库未加载，无法进行转换。请检查网络连接并刷新页面。', 'error');
        checkLibraries();
        return;
    }
    
    const htmlContent = elements.htmlInput ? elements.htmlInput.value.trim() : '';
    if (!htmlContent) {
        showNotification('请输入HTML代码！', 'error');
        if(elements.htmlInput) elements.htmlInput.focus();
        return;
    }
    
    // 验证HTML
    if (!validateHTML(htmlContent)) {
        showNotification('HTML格式似乎有问题，是否继续？', 'warning');
    }
    
    isConverting = true;
    
    // UI状态更新
    if (elements.convertBtn) elements.convertBtn.disabled = true;
    if (elements.btnSpinner) elements.btnSpinner.style.display = 'inline-block';
    
    // 重置状态
    resetConversionStatus();
    
    try {
        // 获取转换选项
        const options = {
            mode: elements.conversionModeSelect ? elements.conversionModeSelect.value : 'hybrid',
            title: elements.pptTitleInput ? elements.pptTitleInput.value : '转换的演示文稿',
            preserveBackground: elements.preserveBackgroundCheckbox ? elements.preserveBackgroundCheckbox.checked : true,
            extractText: elements.extractTextCheckbox ? elements.extractTextCheckbox.checked : true,
        };
        
        // 创建转换器实例
        currentConverter = new HtmlToPptConverter(htmlContent, options);
        
        // 注册状态更新回调
        currentConverter.onStatusUpdate((step, status, message) => {
            updateConversionStatus(step, status, message);
        });
        
        // 开始转换
        const startTime = performance.now();
        const result = await currentConverter.convert();
        const endTime = performance.now();
        
        // 处理结果
        if (result.success && result.blob) {
            const duration = ((endTime - startTime) / 1000).toFixed(2);
            showConversionSuccess(result, duration);
            
            // 保存转换历史
            saveConversionHistory({
                title: options.title,
                mode: options.mode,
                date: new Date().toISOString(),
                duration: duration
            });
        } else {
            showConversionError(result.message);
        }
        
    } catch (error) {
        console.error('转换过程出错:', error);
        showConversionError(error.message);
    } finally {
        // 恢复UI状态
        isConverting = false;
        if (elements.convertBtn) elements.convertBtn.disabled = false;
        if (elements.btnSpinner) elements.btnSpinner.style.display = 'none';
        
        // 清理资源
        if (currentConverter) {
            currentConverter.cleanup();
            currentConverter = null;
        }
    }
}

/**
 * 验证HTML格式
 */
function validateHTML(html) {
    try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        const errorNode = doc.querySelector('parsererror');
        return !errorNode;
    } catch (error) {
        return false;
    }
}

/**
 * 重置转换状态
 */
function resetConversionStatus() {
    const steps = ['parse', 'render', 'extract', 'generate'];
    steps.forEach(step => {
        const statusEl = document.getElementById(`status-${step}`);
        if (statusEl) {
            const icon = statusEl.querySelector('.status-icon');
            icon.className = 'status-icon status-pending';
            icon.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle></svg>';
        }
    });
    
    if (elements.conversionResult) {
        elements.conversionResult.style.display = 'none';
        elements.conversionResult.innerHTML = '';
    }
}

/**
 * 更新转换状态
 */
function updateConversionStatus(step, status, message) {
    const statusEl = document.getElementById(`status-${step}`);
    if (!statusEl) return;
    
    const icon = statusEl.querySelector('.status-icon');
    const text = statusEl.querySelector('.status-text');
    
    // 更新图标和样式
    icon.className = `status-icon status-${status}`;
    
    switch(status) {
        case 'processing':
            icon.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><path d="M12 6v6l4 2"></path></svg>';
            break;
        case 'success':
            icon.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>';
            break;
        case 'error':
            icon.innerHTML = '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>';
            break;
    }
    
    // 更新文本
    if (message && text) {
        text.textContent = message;
    }
}

/**
 * 显示转换成功
 */
function showConversionSuccess(result, duration) {
    elements.conversionResult.innerHTML = `
        <div class="success-message">
            <h4>✅ 转换成功！</h4>
            <p>文件已自动下载</p>
            <div class="conversion-details">
                <p><strong>文件名:</strong> ${result.details?.fileName || elements.pptTitleInput.value + '.pptx'}</p>
                <p><strong>转换模式:</strong> ${getModeName(result.details?.mode)}</p>
                <p><strong>转换时间:</strong> ${duration}秒</p>
            </div>
        </div>
    `;
    elements.conversionResult.classList.add('show', 'slide-up');
    elements.conversionResult.style.display = 'block';
    
    showNotification('PPT文件生成成功！', 'success');
}

/**
 * 显示转换错误
 */
function showConversionError(message) {
    elements.conversionResult.innerHTML = `
        <div class="error-message">
            <h4>❌ 转换失败</h4>
            <p>${message}</p>
            <p class="error-hint">请检查HTML代码是否正确，或尝试其他转换模式。</p>
        </div>
    `;
    elements.conversionResult.classList.add('show', 'slide-up');
    elements.conversionResult.style.display = 'block';
    
    showNotification('转换失败: ' + message, 'error');
}

/**
 * 预览HTML
 */
function previewHTML() {
    const htmlInput = elements.htmlInput.value.trim();
    if (!htmlInput) {
        showNotification('请输入HTML代码！', 'error');
        elements.htmlInput.focus();
        return;
    }
    
    elements.previewContainer.style.display = 'grid';
    elements.previewContainer.classList.add('fade-in');
    
    // 设置iframe内容
    const iframe = elements.htmlPreview;
    iframe.srcdoc = htmlInput;
    
    // 显示预览信息
    elements.conversionInfo.innerHTML = `
        <div class="preview-info">
            <h4>HTML信息</h4>
            <p><strong>大小:</strong> ${formatBytes(new Blob([htmlInput]).size)}</p>
            <p><strong>字符数:</strong> ${htmlInput.length.toLocaleString()}</p>
            <p><strong>建议模式:</strong> ${suggestConversionMode(htmlInput)}</p>
        </div>
    `;
}

/**
 * 关闭预览
 */
function closePreview() {
    elements.previewContainer.style.display = 'none';
}

/**
 * 加载示例
 */
function loadExample() {
    const exampleHTML = getExampleHTML();
    elements.htmlInput.value = exampleHTML;
    autoResizeTextarea();
    
    showNotification('示例已加载！这是一个智慧物流园区的演示页面。', 'success');
    
    // 自动设置标题
    elements.pptTitleInput.value = '智慧物流园区设计方案';
}

/**
 * 显示通知
 */
function showNotification(message, type = 'info') {
    if (!elements.notificationArea) return;

    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.textContent = message;

    elements.notificationArea.appendChild(notification);

    // 动画结束后移除通知，防止DOM堆积
    notification.addEventListener('animationend', (e) => {
        if (e.animationName === 'fadeOut') {
            notification.remove();
        }
    });
}

/**
 * 获取通知图标
 */
function getNotificationIcon(type) {
    const icons = {
        success: '✅',
        error: '❌',
        warning: '⚠️',
        info: 'ℹ️'
    };
    return icons[type] || icons.info;
}

/**
 * 获取模式名称
 */
function getModeName(mode) {
    const modeNames = {
        hybrid: '混合模式',
        screenshot: '截图模式',
        elements: '元素解析模式'
    };
    return modeNames[mode] || mode;
}

/**
 * 格式化字节大小
 */
function formatBytes(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * 建议转换模式
 */
function suggestConversionMode(html) {
    // 简单的启发式判断
    const hasComplexCSS = html.includes('gradient') || html.includes('animation') || html.includes('transform');
    const hasImages = html.includes('<img') || html.includes('background-image');
    const textLength = html.replace(/<[^>]*>/g, '').length;
    
    if (hasComplexCSS || hasImages) {
        return '混合模式（推荐）- 保持视觉效果同时支持文字编辑';
    } else if (textLength > 1000) {
        return '元素解析模式 - 大量文本内容，便于编辑';
    } else {
        return '截图模式 - 简单内容，完美还原';
    }
}

/**
 * 保存转换历史
 */
function saveConversionHistory(record) {
    try {
        const history = JSON.parse(localStorage.getItem('html2ppt_history') || '[]');
        history.unshift(record);
        
        // 只保留最近20条记录
        if (history.length > 20) {
            history.length = 20;
        }
        
        localStorage.setItem('html2ppt_history', JSON.stringify(history));
    } catch (error) {
        console.error('保存历史记录失败:', error);
    }
}

/**
 * 获取示例HTML
 */
function getExampleHTML() {
    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>智慧物流园区设计方案</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: "Microsoft YaHei", Arial, sans-serif;
        }
        
        body {
            background-color: #f5f5f5;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
        }
        
        .container {
            width: 100%;
            max-width: 1200px;
            position: relative;
            background-color: #fff;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15);
        }
        
        .aspect-ratio-box {
            position: relative;
            width: 100%;
            padding-top: 56.25%; /* 16:9 */
        }
        
        .content {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: linear-gradient(135deg, #1a2634 0%, #263545 100%);
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            padding: 5%;
        }
        
        .text-content {
            text-align: center;
        }
        
        .main-title {
            font-size: 3em;
            margin-bottom: 20px;
        }
        
        .sub-title {
            font-size: 1.5em;
            opacity: 0.8;
            margin-bottom: 40px;
        }
        
        .presenter-info {
            font-size: 1.1em;
            opacity: 0.6;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="aspect-ratio-box">
            <div class="content">
                <div class="text-content">
                    <h1 class="main-title">智慧物流园区</h1>
                    <h2 class="sub-title">设计方案与实施路径</h2>
                    <div class="presenter-info">
                        <p>演讲者: Nanford</p>
                        <p>2025/05/27</p>
                    </div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>`;
}

// 导出全局函数
window.convertToPPT = convertToPPT;
window.previewHTML = previewHTML;
window.closePreview = closePreview;
window.loadExample = loadExample;