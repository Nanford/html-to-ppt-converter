// src/js/utils.js

/**
 * 工具函数集合
 * 提供通用的辅助功能
 */

// 添加CSS样式到通知
const notificationStyles = `
.notification {
    position: fixed;
    top: 20px;
    right: 20px;
    max-width: 400px;
    padding: 16px 24px;
    background: white;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
    z-index: 10000;
    display: flex;
    align-items: center;
    gap: 12px;
}

.notification-success {
    border-left: 4px solid #48bb78;
}

.notification-error {
    border-left: 4px solid #f56565;
}

.notification-warning {
    border-left: 4px solid #f6ad55;
}

.notification-info {
    border-left: 4px solid #4299e1;
}

.notification-content {
    display: flex;
    align-items: center;
    gap: 8px;
}

.notification-icon {
    font-size: 1.2em;
}

.notification-message {
    font-size: 0.95em;
    color: #2d3748;
}

.fade-out {
    opacity: 0;
    transform: translateX(20px);
    transition: all 0.3s ease-out;
}
`;

// 注入通知样式
(function injectNotificationStyles() {
    const style = document.createElement('style');
    style.textContent = notificationStyles;
    document.head.appendChild(style);
})();

/**
 * 防抖函数
 */
function debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
        const later = () => {
            clearTimeout(timeout);
            func(...args);
        };
        clearTimeout(timeout);
        timeout = setTimeout(later, wait);
    };
}

/**
 * 节流函数
 */
function throttle(func, limit) {
    let inThrottle;
    return function(...args) {
        if (!inThrottle) {
            func.apply(this, args);
            inThrottle = true;
            setTimeout(() => inThrottle = false, limit);
        }
    };
}

/**
 * 深拷贝对象
 */
function deepClone(obj) {
    if (obj === null || typeof obj !== 'object') return obj;
    if (obj instanceof Date) return new Date(obj.getTime());
    if (obj instanceof Array) return obj.map(item => deepClone(item));
    
    const clonedObj = {};
    for (let key in obj) {
        if (obj.hasOwnProperty(key)) {
            clonedObj[key] = deepClone(obj[key]);
        }
    }
    return clonedObj;
}

/**
 * 生成唯一ID
 */
function generateId() {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * 下载文件
 */
function downloadFile(data, filename, type = 'text/plain') {
    const blob = new Blob([data], { type });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}

/**
 * 复制文本到剪贴板
 */
async function copyToClipboard(text) {
    try {
        if (navigator.clipboard && window.isSecureContext) {
            await navigator.clipboard.writeText(text);
            return true;
        } else {
            // 降级方案
            const textArea = document.createElement('textarea');
            textArea.value = text;
            textArea.style.position = 'fixed';
            textArea.style.left = '-999999px';
            document.body.appendChild(textArea);
            textArea.focus();
            textArea.select();
            
            try {
                document.execCommand('copy');
                return true;
            } catch (error) {
                console.error('复制失败:', error);
                return false;
            } finally {
                document.body.removeChild(textArea);
            }
        }
    } catch (error) {
        console.error('复制到剪贴板失败:', error);
        return false;
    }
}

/**
 * 获取文件扩展名
 */
function getFileExtension(filename) {
    return filename.slice((filename.lastIndexOf('.') - 1 >>> 0) + 2);
}

/**
 * 验证文件类型
 */
function validateFileType(file, allowedTypes) {
    const fileType = file.type || '';
    const fileName = file.name || '';
    const extension = getFileExtension(fileName).toLowerCase();
    
    return allowedTypes.some(type => {
        if (type.startsWith('.')) {
            return extension === type.substring(1);
        }
        return fileType.startsWith(type);
    });
}

/**
 * 格式化日期
 */
function formatDate(date, format = 'YYYY-MM-DD HH:mm:ss') {
    const d = new Date(date);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    const seconds = String(d.getSeconds()).padStart(2, '0');
    
    return format
        .replace('YYYY', year)
        .replace('MM', month)
        .replace('DD', day)
        .replace('HH', hours)
        .replace('mm', minutes)
        .replace('ss', seconds);
}

/**
 * 解析查询字符串
 */
function parseQueryString(queryString) {
    const params = {};
    const searchParams = new URLSearchParams(queryString);
    
    for (const [key, value] of searchParams) {
        params[key] = value;
    }
    
    return params;
}

/**
 * 构建查询字符串
 */
function buildQueryString(params) {
    const searchParams = new URLSearchParams();
    
    for (const [key, value] of Object.entries(params)) {
        if (value !== null && value !== undefined && value !== '') {
            searchParams.append(key, value);
        }
    }
    
    return searchParams.toString();
}

/**
 * 检测浏览器类型
 */
function detectBrowser() {
    const userAgent = navigator.userAgent;
    
    if (userAgent.indexOf('Chrome') > -1) return 'Chrome';
    if (userAgent.indexOf('Safari') > -1) return 'Safari';
    if (userAgent.indexOf('Firefox') > -1) return 'Firefox';
    if (userAgent.indexOf('MSIE') > -1 || userAgent.indexOf('Trident/') > -1) return 'IE';
    if (userAgent.indexOf('Edge') > -1) return 'Edge';
    
    return 'Unknown';
}

/**
 * 检测是否移动设备
 */
function isMobileDevice() {
    return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
}

/**
 * 加载外部脚本
 */
function loadScript(src, callback) {
    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    
    script.onload = () => {
        if (callback) callback(null, script);
    };
    
    script.onerror = () => {
        if (callback) callback(new Error(`Failed to load script: ${src}`));
    };
    
    document.head.appendChild(script);
    return script;
}

/**
 * 延迟执行
 */
function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * 重试函数
 */
async function retry(fn, retries = 3, delayMs = 1000) {
    for (let i = 0; i < retries; i++) {
        try {
            return await fn();
        } catch (error) {
            if (i === retries - 1) throw error;
            await delay(delayMs * Math.pow(2, i)); // 指数退避
        }
    }
}

/**
 * 测量函数执行时间
 */
async function measureTime(fn, label = 'Function') {
    const start = performance.now();
    const result = await fn();
    const end = performance.now();
    console.log(`${label} took ${(end - start).toFixed(2)}ms`);
    return result;
}

/**
 * 简单的事件发射器
 */
class EventEmitter {
    constructor() {
        this.events = {};
    }
    
    on(event, callback) {
        if (!this.events[event]) {
            this.events[event] = [];
        }
        this.events[event].push(callback);
    }
    
    off(event, callback) {
        if (!this.events[event]) return;
        this.events[event] = this.events[event].filter(cb => cb !== callback);
    }
    
    emit(event, ...args) {
        if (!this.events[event]) return;
        this.events[event].forEach(callback => callback(...args));
    }
    
    once(event, callback) {
        const onceWrapper = (...args) => {
            callback(...args);
            this.off(event, onceWrapper);
        };
        this.on(event, onceWrapper);
    }
}

/**
 * 简单的状态管理器
 */
class StateManager {
    constructor(initialState = {}) {
        this.state = initialState;
        this.listeners = [];
    }
    
    getState() {
        return deepClone(this.state);
    }
    
    setState(updates) {
        const oldState = this.state;
        this.state = { ...this.state, ...updates };
        this.notify(oldState, this.state);
    }
    
    subscribe(listener) {
        this.listeners.push(listener);
        return () => {
            this.listeners = this.listeners.filter(l => l !== listener);
        };
    }
    
    notify(oldState, newState) {
        this.listeners.forEach(listener => listener(newState, oldState));
    }
}

/**
 * 颜色转换工具
 */
const ColorUtils = {
    /**
     * RGB颜色转十六进制
     * @param {string} rgb - RGB颜色值
     * @returns {string} 十六进制颜色值
     */
    rgbToHex(rgb) {
        try {
            if (!rgb) return '000000';
            
            // 如果已经是十六进制格式
            if (rgb.startsWith('#')) {
                return rgb.substring(1).padEnd(6, '0');
            }
            
            // 处理rgb()或rgba()格式
            const match = rgb.match(/\d+/g);
            if (!match || match.length < 3) return '000000';
            
            return match.slice(0, 3)
                .map(x => {
                    const num = parseInt(x);
                    return Math.min(255, Math.max(0, num)).toString(16).padStart(2, '0');
                })
                .join('');
        } catch (error) {
            console.warn('颜色转换失败:', rgb, error.message);
            return '000000';
        }
    },
    
    /**
     * 十六进制转RGB
     * @param {string} hex - 十六进制颜色值
     * @returns {object} RGB对象
     */
    hexToRgb(hex) {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result ? {
            r: parseInt(result[1], 16),
            g: parseInt(result[2], 16),
            b: parseInt(result[3], 16)
        } : null;
    },
    
    /**
     * 获取对比色
     * @param {string} color - 背景颜色
     * @returns {string} 对比色
     */
    getContrastColor(color) {
        const rgb = this.hexToRgb(color);
        if (!rgb) return 'FFFFFF';
        
        const brightness = (rgb.r * 299 + rgb.g * 587 + rgb.b * 114) / 1000;
        return brightness > 128 ? '000000' : 'FFFFFF';
    }
};

/**
 * 字体工具
 */
const FontUtils = {
    /**
     * 获取标准字体名称
     * @param {string} fontFamily - CSS字体族
     * @returns {string} 标准字体名称
     */
    getFontFamily(fontFamily) {
        try {
            if (!fontFamily) return 'Calibri';
            
            const family = fontFamily.toLowerCase();
            
            // 中文字体映射
            if (family.includes('yahei') || family.includes('微软雅黑')) return 'Microsoft YaHei';
            if (family.includes('simsun') || family.includes('宋体')) return 'SimSun';
            if (family.includes('simhei') || family.includes('黑体')) return 'SimHei';
            if (family.includes('kaiti') || family.includes('楷体')) return 'KaiTi';
            if (family.includes('fangsong') || family.includes('仿宋')) return 'FangSong';
            
            // 英文字体映射
            if (family.includes('arial')) return 'Arial';
            if (family.includes('helvetica')) return 'Helvetica';
            if (family.includes('times')) return 'Times New Roman';
            if (family.includes('georgia')) return 'Georgia';
            if (family.includes('courier')) return 'Courier New';
            if (family.includes('verdana')) return 'Verdana';
            if (family.includes('tahoma')) return 'Tahoma';
            
            return 'Calibri';
        } catch (error) {
            console.warn('字体解析失败:', fontFamily, error.message);
            return 'Calibri';
        }
    },
    
    /**
     * CSS字体大小转PPT点数
     * @param {string} fontSize - CSS字体大小
     * @returns {number} PPT字体点数
     */
    cssFontSizeToPt(fontSize) {
        try {
            const px = parseFloat(fontSize);
            if (isNaN(px)) return 16;
            
            // 1px ≈ 0.75pt
            return Math.max(8, Math.round(px * 0.75));
        } catch (error) {
            return 16;
        }
    },
    
    /**
     * 判断是否为粗体
     * @param {string} fontWeight - CSS字体粗细
     * @returns {boolean} 是否为粗体
     */
    isBold(fontWeight) {
        if (!fontWeight) return false;
        const weight = fontWeight.toString().toLowerCase();
        return weight === 'bold' || parseInt(weight) >= 600;
    }
};

/**
 * DOM工具
 */
const DOMUtils = {
    /**
     * 检查元素是否可见
     * @param {Element} element - DOM元素
     * @returns {boolean} 是否可见
     */
    isElementVisible(element) {
        try {
            const styles = window.getComputedStyle(element);
            const rect = element.getBoundingClientRect();
            
            return styles.display !== 'none' && 
                   styles.visibility !== 'hidden' && 
                   styles.opacity !== '0' &&
                   rect.width > 0 && 
                   rect.height > 0;
        } catch (error) {
            return false;
        }
    },
    
    /**
     * 获取元素的有效文本内容
     * @param {Element} element - DOM元素
     * @returns {string} 文本内容
     */
    getTextContent(element) {
        try {
            return element.textContent.trim();
        } catch (error) {
            return '';
        }
    },
    
    /**
     * 等待DOM渲染完成
     * @param {number} delay - 延迟时间（毫秒）
     * @returns {Promise} Promise对象
     */
    waitForRender(delay = 1000) {
        return new Promise(resolve => setTimeout(resolve, delay));
    },
    
    /**
     * 查找内容容器
     * @param {Element} root - 根元素
     * @returns {Element} 内容容器
     */
    findContentContainer(root) {
        const selectors = [
            '.content',
            '.aspect-ratio-box',
            '.container',
            '.main',
            '.slide',
            '.page'
        ];
        
        for (const selector of selectors) {
            const element = root.querySelector(selector);
            if (element) {
                console.log('找到内容容器:', selector);
                return element;
            }
        }
        
        // 如果没找到特定容器，返回第一个子元素或根元素
        return root.firstElementChild || root;
    }
};

/**
 * 文本对齐工具
 */
const AlignUtils = {
    /**
     * CSS文本对齐转PPT对齐
     * @param {string} align - CSS文本对齐
     * @returns {string} PPT对齐方式
     */
    getTextAlign(align) {
        switch(align) {
            case 'center': return 'center';
            case 'right': return 'right';
            case 'justify': return 'justify';
            default: return 'left';
        }
    },
    
    /**
     * CSS垂直对齐转PPT垂直对齐
     * @param {string} valign - CSS垂直对齐
     * @returns {string} PPT垂直对齐方式
     */
    getVerticalAlign(valign) {
        switch(valign) {
            case 'middle': return 'middle';
            case 'bottom': return 'bottom';
            default: return 'top';
        }
    }
};

/**
 * 尺寸计算工具
 */
const SizeUtils = {
    /**
     * 像素转英寸
     * @param {number} pixels - 像素值
     * @param {number} dpi - DPI值
     * @returns {number} 英寸值
     */
    pixelsToInches(pixels, dpi = 96) {
        return pixels / dpi;
    },
    
    /**
     * 英寸转像素
     * @param {number} inches - 英寸值
     * @param {number} dpi - DPI值
     * @returns {number} 像素值
     */
    inchesToPixels(inches, dpi = 96) {
        return inches * dpi;
    },
    
    /**
     * 计算相对位置
     * @param {DOMRect} elementRect - 元素矩形
     * @param {DOMRect} containerRect - 容器矩形
     * @returns {object} 相对位置对象
     */
    calculateRelativePosition(elementRect, containerRect) {
        return {
            x: (elementRect.left - containerRect.left) / containerRect.width,
            y: (elementRect.top - containerRect.top) / containerRect.height,
            width: elementRect.width / containerRect.width,
            height: elementRect.height / containerRect.height
        };
    },
    
    /**
     * 转换为PPT坐标
     * @param {object} relative - 相对位置
     * @param {number} slideWidth - 幻灯片宽度
     * @param {number} slideHeight - 幻灯片高度
     * @returns {object} PPT坐标
     */
    toPptCoordinates(relative, slideWidth = 10, slideHeight = 5.625) {
        return {
            x: relative.x * slideWidth,
            y: relative.y * slideHeight,
            w: Math.max(0.1, relative.width * slideWidth),
            h: Math.max(0.1, relative.height * slideHeight)
        };
    }
};

/**
 * 文件工具
 */
const FileUtils = {
    /**
     * 读取文件内容
     * @param {File} file - 文件对象
     * @returns {Promise<string>} 文件内容
     */
    readFileAsText(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(new Error('文件读取失败'));
            reader.readAsText(file, 'UTF-8');
        });
    },
    
    /**
     * 生成时间戳文件名
     * @param {string} prefix - 文件名前缀
     * @param {string} extension - 文件扩展名
     * @returns {string} 文件名
     */
    generateTimestampFilename(prefix = 'file', extension = 'pptx') {
        const now = new Date();
        const timestamp = now.toISOString().slice(0, 19).replace(/[:-]/g, '');
        return `${prefix}_${timestamp}.${extension}`;
    }
};

/**
 * 调试工具
 */
const DebugUtils = {
    /**
     * 记录元素信息
     * @param {Element} element - DOM元素
     * @param {string} label - 标签
     */
    logElementInfo(element, label = 'Element') {
        if (!element) {
            console.log(`${label}: null`);
            return;
        }
        
        const rect = element.getBoundingClientRect();
        const styles = window.getComputedStyle(element);
        
        console.log(`${label}:`, {
            tagName: element.tagName,
            className: element.className,
            id: element.id,
            textContent: element.textContent?.substring(0, 50) + '...',
            rect: {
                x: rect.x,
                y: rect.y,
                width: rect.width,
                height: rect.height
            },
            styles: {
                display: styles.display,
                visibility: styles.visibility,
                opacity: styles.opacity,
                fontSize: styles.fontSize,
                fontFamily: styles.fontFamily,
                color: styles.color,
                backgroundColor: styles.backgroundColor
            }
        });
    },
    
    /**
     * 记录转换步骤
     * @param {string} step - 步骤名称
     * @param {any} data - 数据
     */
    logConversionStep(step, data) {
        console.log(`[转换步骤] ${step}:`, data);
    }
};

// 导出工具函数
window.Utils = {
    debounce,
    throttle,
    deepClone,
    generateId,
    downloadFile,
    copyToClipboard,
    getFileExtension,
    validateFileType,
    formatDate,
    parseQueryString,
    buildQueryString,
    detectBrowser,
    isMobileDevice,
    loadScript,
    delay,
    retry,
    measureTime,
    EventEmitter,
    StateManager
};

// 导出工具对象
window.ColorUtils = ColorUtils;
window.FontUtils = FontUtils;
window.DOMUtils = DOMUtils;
window.AlignUtils = AlignUtils;
window.SizeUtils = SizeUtils;
window.FileUtils = FileUtils;
window.DebugUtils = DebugUtils;