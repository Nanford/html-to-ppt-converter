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
                    fileName: this.options.title + '.pptx'
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

            // 如果启用了文字提取选项，使用分阶段处理
            if (this.options.extractText) {
                console.log('截图模式 + 文字提取：采用分阶段处理...');
                
                // ===== 第一阶段：生成背景图 =====
                this.updateStatus('extract', 'processing', '第一阶段：生成背景图片...');
                console.log('🎨 第一阶段：生成背景图片...');
                
                // 保存文字内容并临时清空
                const originalTextContents = this.temporarilyHideTextContent(captureTarget);

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
                console.log('✅ 第一阶段完成：背景图片已添加');

                // 恢复文字内容
                this.restoreTextContent(originalTextContents);
                console.log('✅ 文字内容已恢复');

                // ===== 第二阶段：等待DOM稳定，然后提取文字 =====
                this.updateStatus('extract', 'processing', '第二阶段：等待DOM稳定...');
                console.log('🔄 第二阶段：等待DOM稳定，准备提取文字...');
                
                await new Promise(resolve => setTimeout(resolve, 1000));
                
                this.updateStatus('extract', 'processing', '第二阶段：提取可编辑文字...');
                console.log('📝 第二阶段：开始文字识别和提取...');
                
                // ===== 第三阶段：文字识别和提取 =====
                await this.addEditableText(slide, captureTarget);
                
                console.log('✅ 截图模式 + 文字提取完成');
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
            
            // 保存原始内容的引用，用于后续恢复
            let originalTextContents = null;
            
            // ===== 第一阶段：生成背景图 =====
            if (this.options.preserveBackground) {
                this.updateStatus('extract', 'processing', '第一阶段：生成背景截图（保留所有视觉元素）...');
                try {
                    console.log('🎨 第一阶段：生成背景截图（只清空文字内容，保留所有其他视觉元素）...');
                    
                    // 保存文字内容并临时清空
                    originalTextContents = this.temporarilyHideTextContent(contentElement);

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
                    console.log('✅ 第一阶段完成：背景截图已添加（包含所有非文字视觉元素）。');
                    
                } catch (bgError) {
                    console.error('混合模式背景截图失败:', bgError);
                    this.updateStatus('extract', 'warning', '背景截图失败，将继续处理文本。');
                }
            } else {
                console.log('混合模式：用户选择不保留背景，跳过背景截图');
                this.updateStatus('extract', 'processing', '跳过背景处理...');
            }
            
            // ===== 第二阶段：恢复文字并等待稳定 =====
            if (this.options.extractText) {
                try {
                    this.updateStatus('extract', 'processing', '第二阶段：恢复文字内容...');
                    console.log('🔄 第二阶段：恢复文字内容...');
                    
                    // 先恢复文字内容（如果之前进行了背景截图的话）
                    if (originalTextContents) {
                        this.restoreTextContent(originalTextContents);
                        console.log('✅ 文字内容已恢复');
                    }
                    
                    // 等待DOM稳定，确保文字已经完全恢复
                    this.updateStatus('extract', 'processing', '第二阶段：等待DOM稳定...');
                    console.log('🔄 第二阶段：等待DOM稳定，准备提取文字...');
                    await new Promise(resolve => setTimeout(resolve, 1500)); // 增加等待时间
                    
                    this.updateStatus('extract', 'processing', '第二阶段：开始提取可编辑文字...');
                    console.log('📝 第二阶段：开始文字识别和提取...');
                    
                    // ===== 第三阶段：文字识别和提取 =====
                    const extractSuccess = await this.addEditableText(slide, contentElement);
                    
                    if (extractSuccess) {
                        console.log('✅ 第二阶段完成：文字提取成功');
                    } else {
                        console.log('⚠️ 第二阶段完成：文字提取失败，但背景已处理');
                    }
                    
                } catch (extractError) {
                    console.warn('混合模式文本提取失败:', extractError.message);
                    this.updateStatus('extract', 'warning', '文本提取失败，但背景已处理完成。');
                }
            } else {
                // 如果不提取文字，也要恢复文字内容（用于后续可能的操作）
                if (originalTextContents) {
                    this.restoreTextContent(originalTextContents);
                    console.log('✅ 文字内容已恢复（虽然不提取文字）');
                }
                console.log('混合模式：用户选择不提取文字，跳过文字处理');
                this.updateStatus('extract', 'processing', '跳过文字提取...');
            }
            
            // ===== 最终阶段：检查处理结果 =====
            if (!this.options.preserveBackground && !this.options.extractText) {
                console.warn('混合模式：用户既不保留背景也不提取文字，生成空白幻灯片');
                this.updateStatus('extract', 'warning', '既未保留背景也未提取文字，生成空白幻灯片');
            } else {
                const resultParts = [];
                if (this.options.preserveBackground) resultParts.push('背景截图（所有视觉元素）');
                if (this.options.extractText) resultParts.push('可编辑文字');
                this.updateStatus('extract', 'success', `混合模式完成：包含${resultParts.join(' + ')}`);
            }
            
            console.log('🎉 混合模式处理完成');
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
            // 保存原始文字内容 - 使用简单的textContent方法
            const originalText = el.textContent;
            originalContents.set(el, originalText);
            
            // 清空文字内容
            el.textContent = '';
        });
        
        console.log(`已临时清空 ${textElements.length} 个元素的文字内容，保留所有视觉样式和装饰元素`);
        return originalContents;
    }
    
    /**
     * 恢复文字内容
     */
    restoreTextContent(originalContents) {
        originalContents.forEach((originalText, el) => {
            // 直接恢复文字内容
            el.textContent = originalText;
        });
        console.log(`已恢复 ${originalContents.size} 个元素的文字内容`);
        
        // 添加调试信息，验证恢复是否成功
        console.log('验证恢复结果：');
        let count = 0;
        originalContents.forEach((originalText, el) => {
            const currentText = el.textContent;
            console.log(`  元素 ${el.tagName}.${el.className}: 原始="${originalText}" 当前="${currentText}" 匹配=${originalText === currentText}`);
            if (originalText === currentText) count++;
        });
        console.log(`恢复成功率: ${count}/${originalContents.size} (${((count/originalContents.size)*100).toFixed(1)}%)`);
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
     * 添加可编辑文字到PPT - 增强版本
     */
    async addEditableText(slide, contentElement) {
        console.log('🔍 开始精确提取文字元素...');
        console.log('内容元素:', contentElement);
                    
        // 重新获取内容区域的位置信息 - 使用可见区域
        const renderArea = document.getElementById('hiddenRenderArea');
        let updatedContentRect;
        
        // 优先使用渲染区域的位置，如果不可用则使用内容元素的位置
        if (renderArea && renderArea.getBoundingClientRect().width > 0) {
            updatedContentRect = renderArea.getBoundingClientRect();
            console.log('📐 使用渲染区域位置:', updatedContentRect);
        } else {
            // 如果内容元素在屏幕外，尝试获取其在渲染时的实际尺寸
            const tempRect = contentElement.getBoundingClientRect();
            if (tempRect.x < -1000 || tempRect.y < -1000) {
                // 元素在屏幕外，使用计算的尺寸
                updatedContentRect = {
                    left: 0,
                    top: 0,
                    width: Math.max(tempRect.width, 1200),
                    height: Math.max(tempRect.height, 675),
                    x: 0,
                    y: 0
                };
                console.log('📐 元素在屏幕外，使用计算尺寸:', updatedContentRect);
            } else {
                updatedContentRect = tempRect;
                console.log('📐 内容区域位置:', updatedContentRect);
            }
        }
        
        // 第一步：收集所有可能的文字元素
        const allTextElements = this.collectAllTextElements(contentElement);
        console.log(`🎯 第一步：找到所有潜在文字元素 ${allTextElements.length} 个`);
        
        // 打印前几个元素的详细信息 - 修复调试输出
        allTextElements.slice(0, 5).forEach((item, index) => {
            const actualText = item.element.textContent?.trim() || '';
            console.log(`  ${index + 1}. ${item.selector} - "${actualText.substring(0, 30)}${actualText.length > 30 ? '...' : ''}"`);
        });
        
        // 第二步：过滤和分析文字元素
        const validTextElements = this.filterValidTextElements(allTextElements);
        console.log(`✅ 第二步：过滤后有效文字元素 ${validTextElements.length} 个`);
        
        // 如果过滤后没有元素，显示被过滤的原因
        if (validTextElements.length === 0 && allTextElements.length > 0) {
            console.warn('⚠️ 所有文字元素都被过滤掉了，检查过滤条件...');
            allTextElements.forEach((item, index) => {
                const el = item.element;
                const actualText = el.textContent?.trim() || '';
                const isVisible = this.isElementVisible(el);
                const hasMeaningful = this.hasMeaningfulTextContent(el);
                const rect = el.getBoundingClientRect();
                const hasSize = rect.width > 0 && rect.height > 0;
                const isNotDuplicate = !this.isNestedDuplicate(el, allTextElements);
                
                console.log(`  ${index + 1}. "${actualText.substring(0, 20)}..." 
                    可见: ${isVisible}, 有意义: ${hasMeaningful}, 有大小: ${hasSize}, 非重复: ${isNotDuplicate}`);
                
                // 详细分析为什么不被认为是有意义的
                if (!hasMeaningful && actualText.length > 0) {
                    console.log(`    文字内容: "${actualText}", 长度: ${actualText.length}, 中文检测: ${/[\u4e00-\u9fff]/.test(actualText)}, 英文数字检测: ${/\w/.test(actualText)}`);
                }
            });
        }
        
        // 第三步：按层级和重要性排序
        const sortedTextElements = this.sortTextElementsByPriority(validTextElements);
        console.log(`📊 第三步：按优先级排序完成`);
        
        // 第四步：精确提取文字并添加到PPT
        let addedCount = 0;
        for (const textItem of sortedTextElements) {
            try {
                const actualText = textItem.element.textContent?.trim() || '';
                console.log(`🔄 尝试提取: "${actualText.substring(0, 20)}..."`);
                const success = await this.extractAndAddTextElement(slide, textItem, contentElement, updatedContentRect);
                if (success) {
                    addedCount++;
                } else {
                    console.warn(`❌ 提取失败: "${actualText.substring(0, 20)}..."`);
                }
            } catch (textError) {
                console.error(`💥 处理文字元素失败: ${textError.message}`);
            }
        }
        
        console.log(`🎉 文字提取完成：成功添加 ${addedCount} 个文字元素到PPT`);
        
        // 如果没有添加任何文字，尝试简化的提取方法
        if (addedCount === 0) {
            console.warn('🔧 没有提取到文字，尝试简化方法...');
            const fallbackCount = await this.fallbackTextExtraction(slide, contentElement, updatedContentRect);
            return fallbackCount > 0;
        }
        
        return addedCount > 0;
    }
    
    /**
     * 检查元素是否包含有意义的文字内容 - 修复中文支持
     */
    hasMeaningfulTextContent(element) {
        const text = element.textContent?.trim();
        
        if (!text) {
            return false;
        }
        
        // 放宽长度限制，支持单个字符（如中文）
        if (text.length < 1) {
            return false;
        }
        
        // 过滤掉只包含特殊字符的内容，但保留中文、英文、数字
        if (/^[\s\-_.,!?;:()[\]{}'"]+$/.test(text)) {
            return false;
        }
        
        // 检查是否包含实际的文字内容（中文、英文、数字）
        // 修复正则表达式，确保正确匹配中文
        const hasChinese = /[\u4e00-\u9fff]/.test(text);
        const hasAlphaNum = /[a-zA-Z0-9]/.test(text);
        
        if (hasChinese || hasAlphaNum) {
            return true;
        }
        
        return false;
    }
    
    /**
     * 检查是否为嵌套重复元素 - 进一步简化逻辑
     */
    isNestedDuplicate(element, allElements) {
        const elementText = element.textContent.trim();
        if (!elementText) return true;
        
        // 对于重要的元素类型，不进行重复检查
        const importantSelectors = ['.main-title', '.sub-title', '.presenter-name', '.presentation-date', 'h1', 'h2', 'h3'];
        const elementClasses = element.className || '';
        const elementTag = element.tagName.toLowerCase();
        
        for (const selector of importantSelectors) {
            if (selector.startsWith('.') && elementClasses.includes(selector.slice(1))) {
                return false; // 重要元素不认为是重复的
            }
            if (selector === elementTag) {
                return false; // 重要元素不认为是重复的
            }
        }
        
        // 简化重复检查：只检查完全相同的文字内容且为直接父子关系
        for (const item of allElements) {
            const otherEl = item.element;
            if (otherEl === element) continue;
            
            // 如果其他元素是当前元素的直接父元素
            if (otherEl === element.parentElement) {
                const otherText = otherEl.textContent.trim();
                // 只有在文字完全相同且长度较短时才认为是重复
                if (otherText === elementText && elementText.length < 20) {
                    return true;
                }
            }
        }
        
        return false;
    }
    
    /**
     * 后备文字提取方法 - 修复错误处理
     */
    async fallbackTextExtraction(slide, contentElement, contentRect) {
        console.log('🔄 启动后备文字提取方法...');
        
        // 使用最简单的选择器
        const simpleSelectors = ['h1', 'h2', 'h3', 'p', '.main-title', '.sub-title'];
        let addedCount = 0;
        
        for (const selector of simpleSelectors) {
            const elements = contentElement.querySelectorAll(selector);
            console.log(`查找选择器 "${selector}": 找到 ${elements.length} 个元素`);
            
            for (const el of elements) {
                const text = el.textContent?.trim();
                if (!text) continue;
                
                console.log(`处理元素: "${text.substring(0, 30)}..."`);
                
                const rect = el.getBoundingClientRect();
                if (rect.width === 0 || rect.height === 0) {
                    console.log('  跳过：元素大小为0');
                    continue;
                }
                
                // 修复位置计算 - 处理元素在屏幕外的情况
                let relativeX, relativeY;
                
                if (contentRect.x < -1000 || contentRect.y < -1000) {
                    // 内容在屏幕外，使用简化的位置计算
                    // 假设元素在容器内的相对位置
                    relativeX = Math.max(0, Math.min(1, (rect.left + 10000) / contentRect.width));
                    relativeY = Math.max(0, Math.min(1, (rect.top + 10000) / contentRect.height));
                    console.log(`  屏幕外元素位置计算: rect(${rect.left}, ${rect.top}), 相对位置(${relativeX.toFixed(2)}, ${relativeY.toFixed(2)})`);
                } else {
                    // 正常位置计算
                    relativeX = (rect.left - contentRect.left) / contentRect.width;
                    relativeY = (rect.top - contentRect.top) / contentRect.height;
                }
                
                // 确保位置在有效范围内
                if (relativeX < 0 || relativeY < 0 || relativeX >= 1 || relativeY >= 1) {
                    console.log(`  跳过：位置超出范围 (${relativeX.toFixed(2)}, ${relativeY.toFixed(2)})`);
                    
                    // 尝试修正位置
                    relativeX = Math.max(0, Math.min(0.9, relativeX));
                    relativeY = Math.max(0, Math.min(0.9, relativeY));
                    console.log(`  修正后位置: (${relativeX.toFixed(2)}, ${relativeY.toFixed(2)})`);
                }
                
                const pptX = relativeX * this.options.slideWidth;
                const pptY = relativeY * this.options.slideHeight;
                const pptW = Math.max(1, (rect.width / contentRect.width) * this.options.slideWidth);
                const pptH = Math.max(0.3, (rect.height / contentRect.height) * this.options.slideHeight);
                
                try {
                    slide.addText(text, {
                        x: pptX,
                        y: pptY,
                        w: pptW,
                        h: pptH,
                        fontSize: this.calculateFallbackFontSize(el),
                        color: '000000',
                        fontFace: 'Calibri',
                        align: 'left',
                        valign: 'top',
                        wrap: true
                    });
                    
                    addedCount++;
                    console.log(`✅ 后备方法成功添加: "${text.substring(0, 20)}..." 位置: (${pptX.toFixed(2)}, ${pptY.toFixed(2)})`);
                } catch (error) {
                    console.error(`❌ 后备方法失败: ${error.message}`);
                }
            }
        }
        
        console.log(`🎉 后备方法完成：成功添加 ${addedCount} 个文字元素`);
        return addedCount;
    }
    
    /**
     * 计算后备方法的字体大小
     */
    calculateFallbackFontSize(element) {
        if (element.classList.contains('main-title') || element.tagName === 'H1') {
            return 36;
        } else if (element.classList.contains('sub-title') || element.tagName === 'H2') {
            return 28;
        } else if (element.tagName === 'H3') {
            return 24;
        } else if (element.classList.contains('presenter-name')) {
            return 18;
        } else if (element.classList.contains('presentation-date')) {
            return 14;
        } else {
            return 16;
        }
    }
    
    /**
     * 收集所有可能的文字元素
     */
    collectAllTextElements(container) {
        const textElements = [];
        const visited = new Set();
        
        // 定义文字元素的选择器，按优先级排序
        const textSelectors = [
            // 高优先级：明确的文字元素
            '.main-title', '.sub-title', '.title', '.heading',
            '.presenter-name', '.presentation-date', '.author',
            'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
            
            // 中优先级：常见文字容器
            'p', '.text', '.content-text', '.description',
            '.label', '.caption', '.item-text',
            
            // 低优先级：可能包含文字的元素
            'span', 'div', 'a', 'button', 'td', 'th', 'li',
            'em', 'strong', 'b', 'i', 'small'
        ];
        
        // 逐个选择器查找元素
        textSelectors.forEach(selector => {
            try {
                const elements = container.querySelectorAll(selector);
                elements.forEach(el => {
                    if (!visited.has(el)) {
                        visited.add(el);
                        textElements.push({
                            element: el,
                            selector: selector,
                            priority: this.getElementPriority(selector)
                        });
                    }
                });
            } catch (error) {
                console.warn(`选择器 "${selector}" 查询失败:`, error);
            }
        });
        
        // 额外检查：寻找直接包含文字但没有被选择器覆盖的元素
        this.findAdditionalTextElements(container, textElements, visited);
        
        return textElements;
    }
    
    /**
     * 获取元素优先级
     */
    getElementPriority(selector) {
        const priorityMap = {
            '.main-title': 10, '.title': 9, '.heading': 9,
            'h1': 8, 'h2': 7, 'h3': 6,
            '.sub-title': 8, '.presenter-name': 7, '.presentation-date': 6,
            'p': 5, '.text': 5, '.content-text': 5,
            '.description': 4, '.label': 4, '.caption': 4,
            'span': 3, 'div': 2, 'a': 3,
            'button': 3, 'em': 3, 'strong': 3
        };
        return priorityMap[selector] || 1;
    }
    
    /**
     * 寻找额外的文字元素
     */
    findAdditionalTextElements(container, existingElements, visited) {
        const walker = document.createTreeWalker(
            container,
            NodeFilter.SHOW_ELEMENT,
            {
                acceptNode: (node) => {
                    if (visited.has(node)) return NodeFilter.FILTER_REJECT;
                    if (!this.hasDirectTextContent(node)) return NodeFilter.FILTER_REJECT;
                    if (!this.hasMeaningfulTextContent(node)) return NodeFilter.FILTER_REJECT;
                    return NodeFilter.FILTER_ACCEPT;
                }
            }
        );
        
        let node;
        while (node = walker.nextNode()) {
            if (!visited.has(node)) {
                visited.add(node);
                existingElements.push({
                    element: node,
                    selector: 'direct-text',
                    priority: 1
                });
            }
        }
    }
    
    /**
     * 过滤有效的文字元素
     */
    filterValidTextElements(textElements) {
        return textElements.filter(item => {
            const el = item.element;
            
            // 基本可见性检查
            if (!this.isElementVisible(el)) return false;
            
            // 检查是否有有意义的文字内容
            if (!this.hasMeaningfulTextContent(el)) return false;
            
            // 检查元素大小
            const rect = el.getBoundingClientRect();
            if (rect.width === 0 || rect.height === 0) return false;
            
            // 避免重复的嵌套元素
            if (this.isNestedDuplicate(el, textElements)) return false;
            
            return true;
        });
    }
    
    /**
     * 按优先级排序文字元素
     */
    sortTextElementsByPriority(textElements) {
        return textElements.sort((a, b) => {
            // 首先按优先级排序（高优先级在前）
            if (a.priority !== b.priority) {
                return b.priority - a.priority;
            }
            
            // 优先级相同时，按位置排序（从上到下，从左到右）
            const aRect = a.element.getBoundingClientRect();
            const bRect = b.element.getBoundingClientRect();
            
            if (Math.abs(aRect.top - bRect.top) > 10) {
                return aRect.top - bRect.top; // 从上到下
            }
            
            return aRect.left - bRect.left; // 从左到右
        });
    }
    
    /**
     * 提取并添加单个文字元素到PPT
     */
    async extractAndAddTextElement(slide, textItem, contentElement, contentRect) {
        const textEl = textItem.element;
        const textContent = textEl.textContent.trim();
        
        if (!textContent) return false;
        
        const rect = textEl.getBoundingClientRect();
        const styles = window.getComputedStyle(textEl);
        
        // 计算相对于内容区域的位置
        const relativeX = (rect.left - contentRect.left) / contentRect.width;
        const relativeY = (rect.top - contentRect.top) / contentRect.height;
        const relativeW = rect.width / contentRect.width;
        const relativeH = rect.height / contentRect.height;
        
        // 确保位置在有效范围内
        if (relativeX < 0 || relativeY < 0 || relativeX >= 1 || relativeY >= 1 || relativeW <= 0 || relativeH <= 0) {
            console.warn(`文字元素位置超出范围: "${textContent.substring(0, 20)}..."`);
            return false;
        }
        
        // 计算PPT中的位置
        const pptX = relativeX * this.options.slideWidth;
        const pptY = relativeY * this.options.slideHeight;
        const pptW = Math.max(0.5, relativeW * this.options.slideWidth);
        const pptH = Math.max(0.2, relativeH * this.options.slideHeight);
        
        // 计算字体大小
        const fontSize = this.calculateOptimalFontSize(textEl, styles, textItem.priority);
        
        // 构建文字选项
        const textOptions = {
            x: pptX,
            y: pptY,
            w: pptW,
            h: pptH,
            fontSize: fontSize,
            color: this.rgbToHex(styles.color || '#000000'),
            bold: this.isBold(styles.fontWeight),
            italic: styles.fontStyle === 'italic',
            fontFace: this.getFontFamily(styles.fontFamily),
            align: this.getTextAlign(styles.textAlign),
            valign: this.getVerticalAlign(textEl, styles),
            wrap: true,
            margin: 0,
            shrinkText: true // 允许自动调整字体大小以适应文本框
        };
        
        // 根据元素类型和优先级进行特殊调整
        this.applySpecialTextFormatting(textOptions, textEl, textItem.priority);
        
        // 添加文字到PPT
        slide.addText(textContent, textOptions);
        
        // 记录详细信息
        console.log(`✓ 添加文字 [${textItem.selector}]: "${textContent.substring(0, 20)}..." 
            位置: (${pptX.toFixed(2)}, ${pptY.toFixed(2)}) 
            大小: ${pptW.toFixed(2)}x${pptH.toFixed(2)} 
            字号: ${fontSize} 
            优先级: ${textItem.priority}`);
        
        return true;
    }
    
    /**
     * 计算最优字体大小
     */
    calculateOptimalFontSize(element, styles, priority) {
        let baseFontSize = parseInt(styles.fontSize) || 16;
        
        // 根据元素类型调整基础字号
        if (element.classList.contains('main-title') || element.tagName === 'H1') {
            return Math.max(36, baseFontSize * 1.5);
        } else if (element.classList.contains('sub-title') || element.tagName === 'H2') {
            return Math.max(28, baseFontSize * 1.2);
        } else if (element.tagName === 'H3') {
            return Math.max(24, baseFontSize * 1.1);
        } else if (element.classList.contains('presenter-name') || element.classList.contains('presentation-date')) {
            return Math.max(14, baseFontSize * 0.9);
        } else if (priority >= 8) {
            // 高优先级元素
            return Math.max(20, baseFontSize);
        } else if (priority >= 5) {
            // 中优先级元素
            return Math.max(16, baseFontSize * 0.95);
        } else {
            // 低优先级元素
            return Math.max(12, baseFontSize * 0.8);
        }
    }
    
    /**
     * 获取垂直对齐方式
     */
    getVerticalAlign(element, styles) {
        // 根据元素类型推断垂直对齐
        if (element.classList.contains('main-title') || element.tagName.match(/^H[1-3]$/)) {
            return 'middle';
        }
        
        const lineHeight = styles.lineHeight;
        if (lineHeight === 'normal' || parseFloat(lineHeight) <= 1.2) {
            return 'middle';
        }
        
        return 'top';
    }
    
    /**
     * 应用特殊文字格式
     */
    applySpecialTextFormatting(textOptions, element, priority) {
        // 主标题特殊格式
        if (element.classList.contains('main-title') || element.tagName === 'H1') {
            textOptions.bold = true;
            textOptions.fontSize = Math.max(textOptions.fontSize, 40);
        }
        
        // 副标题特殊格式
        else if (element.classList.contains('sub-title') || element.tagName === 'H2') {
            textOptions.fontSize = Math.max(textOptions.fontSize, 24);
        }
        
        // 演讲者信息特殊格式
        else if (element.classList.contains('presenter-name')) {
            textOptions.fontSize = Math.max(textOptions.fontSize, 16);
            textOptions.bold = true;
        }
        
        // 日期信息特殊格式
        else if (element.classList.contains('presentation-date')) {
            textOptions.fontSize = Math.max(textOptions.fontSize, 14);
            if (textOptions.color === '000000') {
                textOptions.color = '666666'; // 使用灰色
            }
        }
        
        // 根据优先级调整
        if (priority >= 8) {
            textOptions.bold = textOptions.bold || true;
        }
    }
    
    /**
     * 检查元素是否直接包含文字内容（不是通过子元素） - 放宽条件
     */
    hasDirectTextContent(element) {
        // 检查是否有直接的文本节点
        for (let child of element.childNodes) {
            if (child.nodeType === Node.TEXT_NODE && child.textContent.trim()) {
                return true;
            }
        }
        
        // 如果没有直接文本节点，但只有一个子元素且该子元素包含文字，也算作有效
        const children = Array.from(element.children);
        if (children.length === 1 && children[0].textContent.trim()) {
            // 检查子元素是否为简单的文字容器（如span, em, strong等）
            const childTag = children[0].tagName.toLowerCase();
            if (['span', 'em', 'strong', 'b', 'i', 'small', 'sub', 'sup'].includes(childTag)) {
                return true;
            }
        }
        
        // 对于特定的元素类型，放宽限制
        const elementTag = element.tagName.toLowerCase();
        const elementClass = element.className || '';
        
        // 标题元素和重要类名的元素，即使通过子元素包含文字也算有效
        if (['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(elementTag) ||
            elementClass.includes('title') || 
            elementClass.includes('heading') ||
            elementClass.includes('presenter') ||
            elementClass.includes('date')) {
            return element.textContent.trim().length > 0;
        }
        
        return false;
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
    
    // 计算合适的字号 - 保持原有方法作为后备
    calculateFontSize(element, styles) {
        // 调用新的优化方法
        return this.calculateOptimalFontSize(element, styles, 5); // 默认中等优先级
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