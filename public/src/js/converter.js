// src/js/converter.js

/**
 * HTML to PPT Converter Core
 * 处理HTML到PPT的转换逻辑
 */

class HtmlToPptConverter {
    constructor(htmlContent, options = {}) {
        this.htmlContent = htmlContent;
        this.options = {
            mode: options.mode || 'hybrid',
            preserveBackground: options.preserveBackground !== false,
            extractText: options.extractText !== false,
            highQuality: options.highQuality || false,
            title: options.title || '转换的演示文稿',
            slideWidth: 10, // 英寸
            slideHeight: 5.625, // 英寸 (16:9)
            dpi: options.highQuality ? 300 : 192
        };
        
        // 检查PptxGenJS是否可用
        if (typeof PptxGenJS === 'undefined') {
            throw new Error('PptxGenJS库未加载，请检查网络连接');
        }
        
        this.pptx = new PptxGenJS();
        this.statusCallbacks = {};
    }
    
    /**
     * 注册状态回调
     */
    onStatusUpdate(callback) {
        this.statusCallbacks.update = callback;
    }
    
    /**
     * 更新转换状态
     */
    updateStatus(step, status, message) {
        if (this.statusCallbacks.update) {
            this.statusCallbacks.update(step, status, message);
        }
    }
    
    /**
     * 主转换方法
     */
    async convert() {
        try {
            console.log('开始转换过程...');
            
            // 更新状态
            this.updateStatus('parse', 'processing', '正在解析HTML结构...');
            
            // 设置PPT属性
            this.pptx.title = this.options.title;
            this.pptx.author = 'HTML to PPT Converter';
            this.pptx.layout = 'LAYOUT_16x9';
            
            console.log('PPT设置完成，开始渲染HTML...');
            
            // 渲染HTML到隐藏区域
            const renderArea = document.getElementById('hiddenRenderArea');
            if (!renderArea) {
                throw new Error('渲染区域未找到');
            }
            
            renderArea.innerHTML = this.htmlContent;
            
            // 等待渲染完成
            await this.waitForRender();
            this.updateStatus('render', 'success', '页面渲染完成');
            
            console.log('HTML渲染完成，查找内容元素...');
            
            // 查找主要内容区域
            const contentElement = renderArea.querySelector('.content') || 
                                 renderArea.querySelector('.aspect-ratio-box') || 
                                 renderArea.querySelector('body') || 
                                 renderArea.firstElementChild || 
                                 renderArea;
            
            if (!contentElement) {
                throw new Error('未找到有效的内容元素');
            }
            
            console.log('找到内容元素，开始转换模式:', this.options.mode);
            
            // 根据模式进行转换
            const slide = this.pptx.addSlide();
            
            switch(this.options.mode) {
                case 'screenshot':
                    await this.screenshotMode(slide, contentElement);
                    break;
                case 'elements':
                    await this.elementsMode(slide, contentElement);
                    break;
                case 'hybrid':
                default:
                    await this.hybridMode(slide, contentElement);
                    break;
            }
            
            this.updateStatus('extract', 'processing', '正在提取页面元素...');
            
            // 生成PPT
            this.updateStatus('generate', 'processing', '正在生成PPT文件...');
            await this.generatePPT();
            this.updateStatus('generate', 'success', 'PPT生成成功');
            
            return { 
                success: true, 
                message: '转换成功！',
                details: {
                    mode: this.options.mode,
                    slideCount: 1,
                    fileSize: 'N/A'
                }
            };
            
        } catch (error) {
            console.error('转换失败:', error);
            const step = this.getCurrentStep();
            this.updateStatus(step, 'error', error.message);
            return { 
                success: false, 
                message: '转换失败: ' + error.message,
                error: error
            };
        }
    }
    
    async waitForRender() {
        return new Promise(resolve => setTimeout(resolve, 1000));
    }
    
    async screenshotMode(slide, element) {
        try {
            this.updateStatus('extract', 'processing', '准备截图...');
            console.log('截图模式：开始截图...');
            if (typeof html2canvas === 'undefined') {
                throw new Error('html2canvas库未加载');
            }

            const { captureTarget, finalWidth, finalHeight } = this._calculateCaptureDetails(element, '截图模式');

            // 如果启用了文字提取选项，使用混合模式的方法
            if (this.options.extractText) {
                console.log('截图模式 + 文字提取：生成背景图片...');
                
                // 保存文字内容并临时清空
                const originalTextContents = this.temporarilyHideTextContent(captureTarget);

                this.updateStatus('extract', 'processing', `正在生成背景图片 ${finalWidth}x${finalHeight}...`);
                const canvas = await html2canvas(captureTarget, {
                    width: finalWidth,
                    height: finalHeight,
                    x: 0,
                    y: 0,
                    scale: 1,
                    useCORS: true,
                    allowTaint: true,
                    logging: this.options.debug || false,
                    backgroundColor: window.getComputedStyle(captureTarget).backgroundColor || '#ffffff',
                    scrollX: -captureTarget.scrollLeft,
                    scrollY: -captureTarget.scrollTop
                });

                const imageData = canvas.toDataURL('image/png');
                this.updateStatus('extract', 'processing', '背景图片完成，添加到PPT...');
                
                slide.addImage({
                    data: imageData,
                    x: 0,
                    y: 0,
                    w: this.options.slideWidth,
                    h: this.options.slideHeight,
                    sizing: {
                        type: 'contain',
                        w: this.options.slideWidth,
                        h: this.options.slideHeight
                    }
                });

                // 恢复文字内容
                this.restoreTextContent(originalTextContents);

                // 等待DOM稳定后提取文字
                await new Promise(resolve => setTimeout(resolve, 300));
                
                this.updateStatus('extract', 'processing', '提取可编辑文字...');
                console.log('截图模式：提取文字元素...');
                
                // 添加可编辑文字
                await this.addEditableText(slide, captureTarget);
                
                console.log('截图模式 + 文字提取完成');
            } else {
                // 传统的纯截图模式
                this.updateStatus('extract', 'processing', `正在截取完整图片 ${finalWidth}x${finalHeight}...`);
                const canvas = await html2canvas(captureTarget, {
                    width: finalWidth,
                    height: finalHeight,
                    x: 0,
                    y: 0,
                    scale: 1,
                    useCORS: true,
                    allowTaint: true,
                    logging: this.options.debug || false,
                    backgroundColor: window.getComputedStyle(captureTarget).backgroundColor || '#ffffff',
                    scrollX: -captureTarget.scrollLeft,
                    scrollY: -captureTarget.scrollTop
                });
            
                const imageData = canvas.toDataURL('image/png');
                this.updateStatus('extract', 'processing', '截图完成，添加到PPT...');
                console.log('传统截图完成，添加到PPT...');
            
                slide.addImage({
                    data: imageData,
                    x: 0,
                    y: 0,
                    w: this.options.slideWidth,
                    h: this.options.slideHeight,
                    sizing: {
                        type: 'contain',
                        w: this.options.slideWidth,
                        h: this.options.slideHeight
                    }
                });
            }
            
            this.updateStatus('extract', 'success');
            console.log('截图模式完成');
        } catch (error) {
            console.error('截图模式失败:', error);
            this.updateStatus('extract', 'error', `截图失败: ${error.message}`);
            throw error;
        }
    }
    
    async elementsMode(slide, element) {
        // 提取背景
        if (this.options.preserveBackground) {
            await this.extractBackground(slide, element);
        }
        
        // 提取所有文本元素
        const textElements = element.querySelectorAll('h1, h2, h3, p, div.item-text, div.presenter-name');
        
        for (const textEl of textElements) {
            const rect = textEl.getBoundingClientRect();
            const styles = window.getComputedStyle(textEl);
            
            if (textEl.textContent.trim()) {
                const options = this.getTextOptions(textEl, rect, styles);
                slide.addText(textEl.textContent.trim(), options);
            }
        }
        
        // 提取SVG图标
        const svgElements = element.querySelectorAll('svg');
        for (const svg of svgElements) {
            await this.extractSvg(slide, svg);
        }
        
        this.updateStatus('extract', 'success');
    }
    
    async hybridMode(slide, element) {
        try {
            this.updateStatus('extract', 'processing', '准备混合模式处理...');
            console.log('混合模式：开始处理...');
            console.log(`混合模式配置: 保留背景=${this.options.preserveBackground}, 提取文字=${this.options.extractText}`);

            const { captureTarget: contentElement, finalWidth: bgFinalWidth, finalHeight: bgFinalHeight } = this._calculateCaptureDetails(element, '混合模式背景');
            
            // 1. 处理背景（除文字外的所有视觉元素）
            if (this.options.preserveBackground) {
                this.updateStatus('extract', 'processing', '正在生成背景截图（保留所有视觉元素）...');
                try {
                    console.log('混合模式：生成背景截图（只清空文字内容，保留所有其他视觉元素）...');
                    
                    // 保存文字内容并临时清空
                    const originalTextContents = this.temporarilyHideTextContent(contentElement);

                    console.log(`混合模式背景截图，目标元素: ${contentElement.tagName}.${contentElement.className}, 计划尺寸: ${bgFinalWidth}x${bgFinalHeight}`);
                    const canvas = await html2canvas(contentElement, {
                        width: bgFinalWidth,
                        height: bgFinalHeight,
                        x: 0,
                        y: 0,
                        scale: 1,
                        useCORS: true,
                        allowTaint: true,
                        logging: this.options.debug || false,
                        backgroundColor: window.getComputedStyle(contentElement).backgroundColor || '#ffffff',
                        scrollX: -contentElement.scrollLeft,
                        scrollY: -contentElement.scrollTop
                    });
                    const backgroundImageData = canvas.toDataURL('image/png');
                    slide.addImage({
                        data: backgroundImageData, x: 0, y: 0,
                        w: this.options.slideWidth, h: this.options.slideHeight,
                        sizing: { type: 'contain', w: this.options.slideWidth, h: this.options.slideHeight }
                    });
                    console.log('混合模式：背景截图已添加（包含所有非文字视觉元素）。');
                    
                    // 恢复文字内容
                    this.restoreTextContent(originalTextContents);
                } catch (bgError) {
                    console.error('混合模式背景截图失败:', bgError);
                    this.updateStatus('extract', 'warning', '背景截图失败，将继续处理文本。');
                }
            } else {
                console.log('混合模式：用户选择不保留背景，跳过背景截图');
                this.updateStatus('extract', 'processing', '跳过背景处理...');
            }
            
            // 2. 处理文字提取
            if (this.options.extractText) {
                try {
                    this.updateStatus('extract', 'processing', '提取可编辑文字...');
                    console.log('混合模式：提取文本元素...');
                    
                    // 等待DOM稳定
                    await new Promise(resolve => setTimeout(resolve, 300));
                    
                    // 添加可编辑文字
                    await this.addEditableText(slide, contentElement);
                    
                } catch (extractError) {
                    console.warn('混合模式文本提取失败:', extractError.message);
                    this.updateStatus('extract', 'warning', '文本提取失败，但背景已处理完成。');
                }
            } else {
                console.log('混合模式：用户选择不提取文字，跳过文字处理');
                this.updateStatus('extract', 'processing', '跳过文字提取...');
            }
            
            // 3. 检查处理结果
            if (!this.options.preserveBackground && !this.options.extractText) {
                console.warn('混合模式：用户既不保留背景也不提取文字，生成空白幻灯片');
                this.updateStatus('extract', 'warning', '既未保留背景也未提取文字，生成空白幻灯片');
            } else {
                const resultParts = [];
                if (this.options.preserveBackground) resultParts.push('背景截图（所有视觉元素）');
                if (this.options.extractText) resultParts.push('可编辑文字');
                this.updateStatus('extract', 'success', `混合模式完成：包含${resultParts.join(' + ')}`);
            }
            
            console.log('混合模式处理完成');
        } catch (error) {
            console.error('混合模式失败:', error);
            this.updateStatus('extract', 'error', `混合模式失败: ${error.message}`);
            throw error;
        }
    }
    
    /**
     * 临时隐藏文字内容但保留所有其他视觉元素
     */
    temporarilyHideTextContent(container) {
        const textElements = this.getTextElementsToProcess(container);
        const originalContents = new Map();
        
        textElements.forEach(el => {
            // 保存原始文字内容
            const textNodes = this.getDirectTextNodes(el);
            const originalTexts = textNodes.map(node => node.textContent);
            originalContents.set(el, originalTexts);
            
            // 清空文字内容，但保留所有其他内容（样式、子元素等）
            textNodes.forEach(node => {
                node.textContent = '';
            });
        });
        
        console.log(`已临时清空 ${textElements.length} 个元素的文字内容，保留所有视觉样式和装饰元素`);
        return originalContents;
    }
    
    /**
     * 恢复文字内容
     */
    restoreTextContent(originalContents) {
        originalContents.forEach((originalTexts, el) => {
            const textNodes = this.getDirectTextNodes(el);
            textNodes.forEach((node, index) => {
                if (index < originalTexts.length) {
                    node.textContent = originalTexts[index];
                }
            });
        });
        console.log(`已恢复 ${originalContents.size} 个元素的文字内容`);
    }
    
    /**
     * 获取元素的直接文本节点
     */
    getDirectTextNodes(element) {
        const textNodes = [];
        const walker = document.createTreeWalker(
            element,
            NodeFilter.SHOW_TEXT,
            {
                acceptNode: function(node) {
                    // 只处理有意义的文本节点
                    const text = node.textContent?.trim();
                    if (!text || text.length < 1) return NodeFilter.FILTER_REJECT;
                    return NodeFilter.FILTER_ACCEPT;
                }
            }
        );
        
        let node;
        while (node = walker.nextNode()) {
            textNodes.push(node);
        }
        
        return textNodes;
    }
    
    /**
     * 添加可编辑文字到PPT
     */
    async addEditableText(slide, contentElement) {
        // 重新获取内容区域的位置信息
        const updatedContentRect = contentElement.getBoundingClientRect();
        console.log('内容区域位置:', updatedContentRect);
        
        // 查找文本元素，使用更精确的选择器
        const textElements = contentElement.querySelectorAll('.main-title, .sub-title, .presenter-name, .presentation-date, h1, h2, h3, p, span, div');
        console.log('找到文本元素数量:', textElements.length);
        
        for (const textEl of textElements) {
            try {
                const textContent = textEl.textContent.trim();
                if (!textContent || !this.hasMeaningfulTextContent(textEl)) continue;
                
                const rect = textEl.getBoundingClientRect();
                const styles = window.getComputedStyle(textEl);
                
                // 检查元素是否可见且包含直接文字内容
                if (!this.isElementVisible(textEl) || rect.width === 0 || rect.height === 0) {
                    continue;
                }
                
                // 只处理直接包含文字的元素，跳过只包含子元素的容器
                if (!this.hasDirectTextContent(textEl)) {
                    continue;
                }
                
                // 计算相对于内容区域的位置
                const relativeX = (rect.left - updatedContentRect.left) / updatedContentRect.width;
                const relativeY = (rect.top - updatedContentRect.top) / updatedContentRect.height;
                const relativeW = rect.width / updatedContentRect.width;
                const relativeH = rect.height / updatedContentRect.height;
                
                // 确保位置在有效范围内
                if (relativeX >= 0 && relativeY >= 0 && relativeX < 1 && relativeY < 1 && relativeW > 0 && relativeH > 0) {
                    
                    // 根据类名调整字号
                    let fontSize = this.calculateFontSize(textEl, styles);
                    
                    // 计算PPT中的位置
                    const pptX = relativeX * this.options.slideWidth;
                    const pptY = relativeY * this.options.slideHeight;
                    const pptW = Math.max(1, relativeW * this.options.slideWidth);
                    const pptH = Math.max(0.3, relativeH * this.options.slideHeight);
                    
                    const textOptions = {
                        x: pptX,
                        y: pptY,
                        w: pptW,
                        h: pptH,
                        fontSize: fontSize,
                        color: this.rgbToHex(styles.color || '#000000'),
                        bold: this.isBold(styles.fontWeight),
                        fontFace: this.getFontFamily(styles.fontFamily),
                        align: this.getTextAlign(styles.textAlign),
                        valign: 'top',
                        wrap: true,
                        margin: 0
                    };
                    
                    // 根据元素类型调整样式
                    if (textEl.classList.contains('main-title') || textEl.tagName === 'H1') {
                        textOptions.fontSize = Math.max(44, fontSize);
                        textOptions.bold = true;
                    } else if (textEl.classList.contains('sub-title') || textEl.tagName === 'H2') {
                        textOptions.fontSize = Math.max(24, fontSize * 0.8);
                    } else if (textEl.classList.contains('presenter-name') || textEl.classList.contains('presentation-date')) {
                        textOptions.fontSize = Math.max(16, fontSize * 0.7);
                    }
                    
                    slide.addText(textContent, textOptions);
                    console.log(`添加文本 "${textContent.substring(0, 20)}..." 位置: (${pptX.toFixed(2)}, ${pptY.toFixed(2)}) 字号: ${textOptions.fontSize}`);
                }
                
            } catch (textError) {
                console.warn('处理文本元素失败:', textError.message);
            }
        }
    }
    
    /**
     * 检查元素是否直接包含文字内容（不是通过子元素）
     */
    hasDirectTextContent(element) {
        for (let child of element.childNodes) {
            if (child.nodeType === Node.TEXT_NODE && child.textContent.trim()) {
                return true;
            }
        }
        return false;
    }
    
    /**
     * 检查元素是否包含有意义的文字内容
     */
    hasMeaningfulTextContent(element) {
        const text = element.textContent?.trim();
        if (!text) return false;
        
        // 过滤掉只包含特殊字符的内容
        if (text.length < 2) return false;
        if (/^[\s\-_.,!?;:()[\]{}'"]+$/.test(text)) return false;
        
        return true;
    }
    
    /**
     * 获取需要处理的文字元素
     */
    getTextElementsToProcess(container) {
        const textElements = [];
        
        // 直接文字元素选择器
        const directTextSelectors = [
            'h1', 'h2', 'h3', 'h4', 'h5', 'h6',  // 标题元素
            'p',                                    // 段落
            'span',                                 // 行内文字
            'a',                                    // 链接
            'em', 'strong', 'b', 'i',              // 强调元素
            'label',                               // 标签
            'button',                              // 按钮
            'td', 'th',                            // 表格单元格
            'li',                                  // 列表项
            'blockquote', 'cite', 'code',          // 引用和代码
            'small', 'sub', 'sup',                 // 其他文字元素
            'div'                                   // div元素也可能包含文字
        ];
        
        // 按类名匹配的文字元素
        const textClassSelectors = [
            '.text', '.title', '.subtitle', '.heading',
            '.content', '.description', '.label', '.caption',
            '.main-title', '.sub-title', '.presenter-name', '.presentation-date',
            '.item-text', '.text-content', '.text-wrapper'
        ];
        
        // 合并所有选择器
        const allSelectors = [...directTextSelectors, ...textClassSelectors];
        
        // 查找所有匹配的元素
        allSelectors.forEach(selector => {
            try {
                const elements = container.querySelectorAll(selector);
                elements.forEach(el => {
                    if (!textElements.includes(el) && 
                        this.isElementVisible(el) && 
                        this.hasDirectTextContent(el)) {
                        textElements.push(el);
                    }
                });
            } catch (error) {
                console.warn(`选择器 "${selector}" 查询失败:`, error);
            }
        });
        
        console.log(`找到需要处理的文字元素数量: ${textElements.length}`);
        
        // 输出详细信息
        textElements.forEach((el, index) => {
            const text = el.textContent?.trim().substring(0, 30) || '';
            console.log(`  ${index + 1}. ${el.tagName}.${el.className || '(无类名)'}: "${text}${text.length > 30 ? '...' : ''}"`);
        });
        
        return textElements;
    }
    
    // 计算合适的字号
    calculateFontSize(element, styles) {
        try {
            let baseFontSize = parseInt(styles.fontSize) || 16;
            
            // 根据元素类型调整基础字号
            if (element.classList.contains('main-title') || element.tagName === 'H1') {
                return Math.max(40, baseFontSize * 1.2);
            } else if (element.classList.contains('sub-title') || element.tagName === 'H2') {
                return Math.max(24, baseFontSize);
            } else if (element.classList.contains('presenter-name') || element.classList.contains('presentation-date')) {
                return Math.max(14, baseFontSize * 0.8);
            } else {
                return Math.max(16, baseFontSize);
            }
        } catch (error) {
            return 16;
        }
    }
    
    // 判断是否为粗体
    isBold(fontWeight) {
        if (!fontWeight) return false;
        const weight = fontWeight.toString().toLowerCase();
        return weight === 'bold' || parseInt(weight) >= 600;
    }
    
    isElementVisible(element) {
        const styles = window.getComputedStyle(element);
        return styles.display !== 'none' && 
               styles.visibility !== 'hidden' && 
               styles.opacity !== '0';
    }
    
    getTextOptions(element, rect, styles) {
        const renderArea = document.getElementById('hiddenRenderArea');
        const renderAreaRect = renderArea.getBoundingClientRect();
        
        return {
            x: ((rect.left - renderAreaRect.left) / this.options.dpi),
            y: ((rect.top - renderAreaRect.top) / this.options.dpi),
            w: (rect.width / this.options.dpi),
            h: (rect.height / this.options.dpi),
            fontSize: parseInt(styles.fontSize),
            color: this.rgbToHex(styles.color),
            bold: styles.fontWeight >= 600,
            fontFace: this.getFontFamily(styles.fontFamily),
            align: this.getTextAlign(styles.textAlign)
        };
    }
    
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
    }
    
    getFontFamily(fontFamily) {
        try {
            if (!fontFamily) return 'Calibri';
            
            const family = fontFamily.toLowerCase();
            if (family.includes('yahei') || family.includes('微软雅黑')) return 'Microsoft YaHei';
            if (family.includes('arial')) return 'Arial';
            if (family.includes('helvetica')) return 'Helvetica';
            if (family.includes('simsun') || family.includes('宋体')) return 'SimSun';
            if (family.includes('simhei') || family.includes('黑体')) return 'SimHei';
            if (family.includes('times')) return 'Times New Roman';
            if (family.includes('georgia')) return 'Georgia';
            
            return 'Calibri';
        } catch (error) {
            console.warn('字体解析失败:', fontFamily, error.message);
            return 'Calibri';
        }
    }
    
    getTextAlign(align) {
        switch(align) {
            case 'center': return 'center';
            case 'right': return 'right';
            case 'justify': return 'justify';
            default: return 'left';
        }
    }
    
    async extractBackground(slide, element) {
        const styles = window.getComputedStyle(element);
        const bgColor = styles.backgroundColor;
        const bgImage = styles.backgroundImage;
        
        if (bgImage && bgImage !== 'none') {
            // 处理渐变背景
            if (bgImage.includes('gradient')) {
                // PPT不支持CSS渐变，使用截图
                const bgCanvas = document.createElement('canvas');
                bgCanvas.width = 1920;
                bgCanvas.height = 1080;
                const ctx = bgCanvas.getContext('2d');
                
                // 这里简化处理，实际可以解析渐变
                ctx.fillStyle = '#1a2634';
                ctx.fillRect(0, 0, 1920, 1080);
                
                slide.background = { data: bgCanvas.toDataURL() };
            }
        } else if (bgColor && bgColor !== 'transparent') {
            slide.background = { color: this.rgbToHex(bgColor) };
        }
    }
    
    async extractSvg(slide, svgElement) {
        const rect = svgElement.getBoundingClientRect();
        const renderAreaRect = document.getElementById('hiddenRenderArea').getBoundingClientRect();
        
        // 将SVG转换为图片
        const svgString = new XMLSerializer().serializeToString(svgElement);
        const svg64 = btoa(svgString);
        const image64 = 'data:image/svg+xml;base64,' + svg64;
        
        slide.addImage({
            data: image64,
            x: ((rect.left - renderAreaRect.left) / this.options.dpi),
            y: ((rect.top - renderAreaRect.top) / this.options.dpi),
            w: (rect.width / this.options.dpi),
            h: (rect.height / this.options.dpi)
        });
    }
    
    /**
     * 生成PPT文件
     */
    async generatePPT() {
        const fileName = this.options.title + '.pptx';
        
        await this.pptx.writeFile({ 
            fileName: fileName,
            compression: this.options.highQuality ? false : true
        });
    }
    
    /**
     * 获取当前步骤
     */
    getCurrentStep() {
        // 简单的步骤判断逻辑
        if (!this.doc) return 'parse';
        if (!this.pptx.slides) return 'render';
        return 'generate';
    }
    
    /**
     * 清理资源
     */
    cleanup() {
        // 清理渲染区域
        const renderArea = document.getElementById('hiddenRenderArea');
        if (renderArea) {
            renderArea.innerHTML = '';
        }
    }

    _calculateCaptureDetails(element, modeName) {
        const contentSelectors = ['.content', '.aspect-ratio-box', '[data-capture-target]', 'main', 'article', '.main-content-area', '.slide-content'];
        let captureTarget = null;
        for (const selector of contentSelectors) {
            const found = element.querySelector(selector);
            if (found) {
                captureTarget = found;
                console.log(`${modeName}: 找到优先选择器 "${selector}" 对应的元素:`, captureTarget);
                break;
            }
        }
        if (!captureTarget) {
            captureTarget = element.matches('body') && element.firstElementChild ? element.firstElementChild : element;
            console.warn(`${modeName}: 未找到优先选择器对应的元素，将使用:`, captureTarget);
        }

        const rect = captureTarget.getBoundingClientRect();
        let w = rect.width;
        let h = rect.height;

        console.log(`${modeName}: 初始 getBoundingClientRect 尺寸: ${w.toFixed(2)}x${h.toFixed(2)} for`, captureTarget);

        if (w === 0 || h === 0) {
            console.warn(`${modeName}: 捕获目标 ${captureTarget.tagName}.${captureTarget.className} getBoundingClientRect 尺寸为零！尝试 clientWidth/Height。`);
            w = captureTarget.clientWidth;
            h = captureTarget.clientHeight;
            if (w === 0 || h === 0) {
                 console.error(`${modeName}: 捕获目标 ${captureTarget.tagName}.${captureTarget.className} 所有尺寸计算都为零！将默认使用1920x1080。`);
                 w = 1920;
                 h = 1080;
            }
        } 
        
        const targetAspectRatio = 16 / 9;
        const currentAspectRatio = w / h;

        if (Math.abs(currentAspectRatio - targetAspectRatio) > 0.01) { // 1% 容差
            console.warn(`${modeName}: 捕获目标 ${captureTarget.tagName}.${captureTarget.className} (${w.toFixed(0)}x${h.toFixed(0)}, ratio ${currentAspectRatio.toFixed(2)}) 不是严格的16:9 (目标 ${targetAspectRatio.toFixed(2)})。将基于宽度调整高度以强制16:9。`);
            h = w / targetAspectRatio;
        }
        
        // 先向上取整
        let calculatedWidth = Math.ceil(w);
        let calculatedHeight = Math.ceil(h);

        // 再增加2px的缓冲区，尝试消除边缘白线
        const buffer = 2; // 增加2px的缓冲区
        const finalWidth = calculatedWidth + buffer;
        const finalHeight = calculatedHeight + buffer;

        console.log(`${modeName}: 计算截图尺寸 (ceil后): ${calculatedWidth}x${calculatedHeight}, 带缓冲区 (+${buffer}px) 后: ${finalWidth}x${finalHeight} for element`, captureTarget);
        return { captureTarget, finalWidth, finalHeight };
    }
}

// 导出到全局
window.HtmlToPptConverter = HtmlToPptConverter;