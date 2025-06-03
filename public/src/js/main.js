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
    html2canvasStatus: null,
    conversionResult: null,
    conversionInfo: null,
    previewContainer: null,
    htmlPreview: null
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
    elements.conversionResult = document.getElementById('conversionResult');
    elements.conversionInfo = document.getElementById('conversionInfo');
    elements.previewContainer = document.getElementById('previewContainer');
    elements.htmlPreview = document.getElementById('htmlPreview');
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
        elements.loadExampleBtn.addEventListener('click', () => {
            const exampleHtml = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>企业数字化转型解决方案</title>
    <style>
        /* 全局样式 */
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: "Microsoft YaHei", "微软雅黑", Arial, sans-serif;
            background-color: #f5f5f5;
            color: #333;
            line-height: 1.6;
        }
        
        /* 主容器 - 16:9 比例 - 优化版本 */
        .main-container {
            width: 1920px;
            height: 1080px;
            margin: 0 auto;
            background: #2c3e50;
            position: relative;
            overflow: hidden;
            box-shadow: 0 0 30px rgba(0, 0, 0, 0.2);
        }
        
        /* 主背景层 */
        .main-background {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 50%, #2c3e50 100%);
        }
        
        /* 纹理图案层 - 优化为更简单的实现 */
        .texture-layer {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-image: 
                linear-gradient(45deg, rgba(255,255,255,0.02) 25%, transparent 25%),
                linear-gradient(-45deg, rgba(255,255,255,0.02) 25%, transparent 25%),
                linear-gradient(45deg, transparent 75%, rgba(255,255,255,0.02) 75%),
                linear-gradient(-45deg, transparent 75%, rgba(255,255,255,0.02) 75%);
            background-size: 20px 20px;
            background-position: 0 0, 0 10px, 10px -10px, -10px 0px;
        }
        
        /* 装饰圆圈 - 使用border和伪元素确保兼容性 */
        .decoration-circle {
            position: absolute;
            border-radius: 50%;
            border: 2px solid rgba(230,126,34,0.3);
        }
        
        .circle-large {
            width: 400px;
            height: 400px;
            top: -200px;
            right: -200px;
        }
        
        .circle-medium {
            width: 250px;
            height: 250px;
            bottom: -125px;
            left: -125px;
            border-color: rgba(255,255,255,0.15);
        }
        
        .circle-small {
            width: 150px;
            height: 150px;
            top: 20%;
            left: 10%;
            border-color: rgba(230,126,34,0.2);
        }
        
        /* 数据点装饰元素 - 确保可见性 */
        .data-dot {
            position: absolute;
            width: 12px;
            height: 12px;
            background: #e67e22;
            border-radius: 50%;
            box-shadow: 0 0 10px rgba(230,126,34,0.5);
        }
        
        .dot-1 { top: 25%; left: 20%; }
        .dot-2 { top: 50%; right: 15%; }
        .dot-3 { bottom: 30%; left: 12%; }
        .dot-4 { top: 70%; right: 25%; }
        .dot-5 { bottom: 15%; right: 40%; }
        
        /* 装饰线条 - 使用CSS绘制 */
        .decoration-line {
            position: absolute;
            background: rgba(255,255,255,0.1);
        }
        
        .line-1 {
            width: 200px;
            height: 1px;
            top: 20%;
            right: 10%;
            transform: rotate(45deg);
        }
        
        .line-2 {
            width: 150px;
            height: 1px;
            bottom: 25%;
            left: 8%;
            transform: rotate(-30deg);
        }
        
        .line-3 {
            width: 100px;
            height: 1px;
            top: 60%;
            left: 5%;
            transform: rotate(60deg);
        }
        
        /* 内容容器 */
        .content-wrapper {
            position: relative;
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            padding: 0 200px;
            text-align: center;
            color: white;
            z-index: 10;
        }
        
        /* 主标题 */
        .main-title {
            font-size: 72px;
            font-weight: 700;
            margin-bottom: 30px;
            letter-spacing: 5px;
            text-shadow: 0 3px 15px rgba(0,0,0,0.4);
            position: relative;
            display: inline-block;
        }
        
        /* 标题装饰线 */
        .main-title::after {
            content: "";
            position: absolute;
            bottom: -20px;
            left: 50%;
            transform: translateX(-50%);
            width: 120px;
            height: 4px;
            background: linear-gradient(90deg, transparent, #e67e22, transparent);
        }
        
        /* 副标题 */
        .sub-title {
            font-size: 32px;
            font-weight: 300;
            margin-bottom: 60px;
            letter-spacing: 2px;
            opacity: 0.9;
            text-shadow: 0 2px 8px rgba(0,0,0,0.3);
        }
        
        /* 公司信息 */
        .company-info {
            position: absolute;
            bottom: 80px;
            left: 0;
            width: 100%;
            text-align: center;
            z-index: 10;
        }
        
        .company-name {
            font-size: 24px;
            font-weight: 400;
            margin-bottom: 10px;
            color: rgba(255,255,255,0.95);
        }
        
        .presentation-date {
            font-size: 18px;
            font-weight: 300;
            color: rgba(255,255,255,0.7);
        }
        
        /* 额外的装饰几何图形 */
        .geometric-shape {
            position: absolute;
            border: 1px solid rgba(255,255,255,0.1);
        }
        
        .triangle-1 {
            width: 0;
            height: 0;
            border-left: 30px solid transparent;
            border-right: 30px solid transparent;
            border-bottom: 40px solid rgba(230,126,34,0.2);
            top: 15%;
            right: 20%;
            border: none;
        }
        
        .square-1 {
            width: 25px;
            height: 25px;
            background: rgba(255,255,255,0.05);
            transform: rotate(45deg);
            top: 70%;
            left: 80%;
        }
        
        .hexagon-1 {
            width: 40px;
            height: 35px;
            background: rgba(230,126,34,0.15);
            position: relative;
            top: 80%;
            right: 30%;
        }
        
        .hexagon-1:before {
            content: "";
            position: absolute;
            top: -10px;
            left: 0;
            width: 0;
            height: 0;
            border-left: 20px solid transparent;
            border-right: 20px solid transparent;
            border-bottom: 10px solid rgba(230,126,34,0.15);
        }
        
        .hexagon-1:after {
            content: "";
            position: absolute;
            bottom: -10px;
            left: 0;
            width: 0;
            height: 0;
            border-left: 20px solid transparent;
            border-right: 20px solid transparent;
            border-top: 10px solid rgba(230,126,34,0.15);
        }
    </style>
</head>
<body>
    <!-- 主容器 - 16:9 比例 -->
    <div class="main-container">
        <!-- 多层背景 -->
        <div class="main-background"></div>
        <div class="texture-layer"></div>
        
        <!-- 装饰圆圈 -->
        <div class="decoration-circle circle-large"></div>
        <div class="decoration-circle circle-medium"></div>
        <div class="decoration-circle circle-small"></div>
        
        <!-- 装饰线条 -->
        <div class="decoration-line line-1"></div>
        <div class="decoration-line line-2"></div>
        <div class="decoration-line line-3"></div>
        
        <!-- 数据点 -->
        <div class="data-dot dot-1"></div>
        <div class="data-dot dot-2"></div>
        <div class="data-dot dot-3"></div>
        <div class="data-dot dot-4"></div>
        <div class="data-dot dot-5"></div>
        
        <!-- 几何装饰形状 -->
        <div class="geometric-shape triangle-1"></div>
        <div class="geometric-shape square-1"></div>
        <div class="geometric-shape hexagon-1"></div>
        
        <!-- 内容区 -->
        <div class="content-wrapper">
            <h1 class="main-title">企业数字化转型解决方案</h1>
            <p class="sub-title">Digital Transformation Solution for Enterprises</p>
        </div>
        
        <!-- 公司信息 -->
        <div class="company-info">
            <div class="company-name">某某科技咨询有限公司</div>
            <div class="presentation-date">2023年11月</div>
        </div>
    </div>
</body>
</html>`;
            
            elements.htmlInput.value = exampleHtml;
            showNotification('已加载包含丰富装饰元素的示例HTML（纹理、线条、几何图形等）', 'success');
        });
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
        if (result.success) {
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
    // 安全检查：确保元素存在
    if (elements.conversionResult) {
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
    } else {
        // 如果没有找到转换结果元素，创建一个临时的成功显示
        console.log('转换成功，但未找到结果显示区域');
        
        // 尝试在页面中找到合适的位置显示成功信息
        const body = document.body;
        const successDiv = document.createElement('div');
        successDiv.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: #fff;
            border: 2px solid #4CAF50;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            z-index: 10000;
            max-width: 400px;
        `;
        successDiv.innerHTML = `
            <h4 style="color: #4CAF50; margin: 0 0 10px 0;">✅ 转换成功！</h4>
            <p style="margin: 0 0 10px 0;">文件已自动下载</p>
            <p style="margin: 0 0 10px 0;"><strong>文件名:</strong> ${result.details?.fileName || elements.pptTitleInput?.value + '.pptx' || '转换的演示文稿.pptx'}</p>
            <p style="margin: 0 0 10px 0;"><strong>转换模式:</strong> ${getModeName(result.details?.mode)}</p>
            <p style="margin: 0 0 15px 0;"><strong>转换时间:</strong> ${duration}秒</p>
            <button onclick="this.parentElement.remove()" style="
                margin-top: 15px;
                padding: 8px 15px;
                background: #4CAF50;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
            ">关闭</button>
        `;
        body.appendChild(successDiv);
        
        // 10秒后自动移除
        setTimeout(() => {
            if (successDiv.parentElement) {
                successDiv.remove();
            }
        }, 10000);
    }
    
    showNotification('PPT文件生成成功！', 'success');
}

/**
 * 显示转换错误
 */
function showConversionError(message) {
    // 安全检查：确保元素存在
    if (elements.conversionResult) {
        elements.conversionResult.innerHTML = `
            <div class="error-message">
                <h4>❌ 转换失败</h4>
                <p>${message}</p>
                <p class="error-hint">请检查HTML代码是否正确，或尝试其他转换模式。</p>
            </div>
        `;
        elements.conversionResult.classList.add('show', 'slide-up');
        elements.conversionResult.style.display = 'block';
    } else {
        // 如果没有找到转换结果元素，创建一个临时的错误显示
        console.error('转换失败:', message);
        
        // 尝试在页面中找到合适的位置显示错误
        const body = document.body;
        const errorDiv = document.createElement('div');
        errorDiv.style.cssText = `
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            background: #fff;
            border: 2px solid #ff4444;
            border-radius: 8px;
            padding: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            z-index: 10000;
            max-width: 400px;
        `;
        errorDiv.innerHTML = `
            <h4 style="color: #ff4444; margin: 0 0 10px 0;">❌ 转换失败</h4>
            <p style="margin: 0 0 10px 0;">${message}</p>
            <p style="margin: 0; color: #666; font-size: 14px;">请检查HTML代码是否正确，或尝试其他转换模式。</p>
            <button onclick="this.parentElement.remove()" style="
                margin-top: 15px;
                padding: 8px 15px;
                background: #ff4444;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
            ">关闭</button>
        `;
        body.appendChild(errorDiv);
        
        // 5秒后自动移除
        setTimeout(() => {
            if (errorDiv.parentElement) {
                errorDiv.remove();
            }
        }, 5000);
    }
    
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