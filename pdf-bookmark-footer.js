/**
 * PDF书签页脚工具
 * 用于在PDF文档中为书签页面添加页脚注释
 * 
 * @copyright Copyright (c) 2025 yhh
 * @license MIT
 * @version 1.0.0
 * @tested Foxit PDF Editor 13.0.1.21693
 */

/**
 * 全局配置对象
 * @property {Object} annotation - 注释相关配置
 * @property {Array} annotation.rect - 页脚区域 [左, 下, 右, 上]
 * @property {number} annotation.textSize - 字体大小
 * @property {string} annotation.font - 字体名称
 * @property {Array} annotation.textColor - 文字颜色 [类型, R, G, B]
 * @property {Array} annotation.strokeColor - 边框颜色 [类型, R, G, B]
 * @property {Array} annotation.fillColor - 背景颜色 [类型, R, G, B]
 */
var CONFIG = {
    annotation: {
        rect: [30, 15, 570, 35],
        textSize: 9,
        font: "SimSun",
        textColor: ["RGB", 0, 0, 0],
        strokeColor: ["RGB", 1, 1, 1],
        fillColor: ["RGB", 1, 1, 1]
    }
};

// 全局变量定义
var gDoc;          // 当前活动的PDF文档
var gBookmarks;    // 文档中的所有书签
var gOriginalPage; // 原始页码（用于恢复）
var gPreviewAnnot; // 预览注释对象

/**
 * 创建页脚注释
 * @param {Object} doc - PDF文档对象
 * @param {number} page - 页码
 * @param {string} text - 注释文本
 * @returns {Object} 创建的注释对象
 * @throws {Error} 创建失败时抛出错误
 */
function createAnnotation(doc, page, text) {
    try {
        if (page < 0 || page >= doc.numPages) {
            throw new Error(`页码 ${page + 1} 超出范围 (总页数: ${doc.numPages})`);
        }

        console.println("  正在创建页脚注释...");
        console.println(`  - 页码: ${page + 1}`);
        console.println("  - 文本: " + text);
        
        var annot = doc.addAnnot({
            // 基本属性
            page: page,
            type: "FreeText",
            intent: "FreeTextTypewriter",
            rect: CONFIG.annotation.rect,
            contents: text,
            alignment: 1,
            
            // 文本样式
            textSize: CONFIG.annotation.textSize,
            richText: true,
            defaultStyle: {
                font: CONFIG.annotation.font,
                textSize: CONFIG.annotation.textSize
            },
            
            // 颜色设置 [类型, R, G, B]
            textColor: CONFIG.annotation.textColor,
            strokeColor: CONFIG.annotation.strokeColor,
            fillColor: CONFIG.annotation.fillColor,
            
            // 外观控制
            opacity: 1,
            rotation: 0,
            AP: "N",
            noView: false,
            
            // 权限控制
            readOnly: true,
            print: true,
            locked: true
        });

        if (annot) {
            annot.setProps({
                strokeWidth: 0,
                borderStyle: "S",
                lineEndingStyle: "None"
            });
            console.println("  ✓ 页脚注释创建成功");
        } else {
            throw new Error("注释创建失败");
        }

        return annot;
    } catch (e) {
        console.println("  ! 创建注释时出错: " + e);
        throw e;
    }
}

/**
 * 显示确认对话框并处理所有书签
 * 主要流程：
 * 1. 确认是否继续处理
 * 2. 遍历处理所有书签
 * 3. 统计处理结果
 * 4. 显示处理报告
 */
function showConfirmDialog() {
    var response = app.alert({
        cMsg: "请查看当前页面的预览效果\n是否继续处理所有书签？",
        cTitle: "确认继续",
        nIcon: 2,
        nType: 2
    });

    if (gPreviewAnnot) {
        gPreviewAnnot.destroy();
    }
    
    if (response !== 4) {
        console.println("用户取消操作");
        return;
    }
    
    console.println("开始处理所有书签...");
    console.println("总书签数量: " + gBookmarks.length);
    
    var results = {success: 0, fail: 0, failedBookmarks: []};
    var errorLog = "";

    // 创建一个数组来跟踪已处理的页码
    var processedPages = new Array(gDoc.numPages).fill(false);
    
    console.println("\n=== 开始处理书签页面 ===");
    for (var i = 0; i < gBookmarks.length; i++) {
        var bookmark = gBookmarks[i];
        try {
            console.println(`\n处理书签 [${i + 1}/${gBookmarks.length}]:`);
            console.println(`书签名称: ${bookmark.name}`);
            
            bookmark.execute();
            var currentPage = gDoc.pageNum;
            processedPages[currentPage] = true;
            
            console.println(`页面编号: ${currentPage + 1}`);
            
            var annot = createAnnotation(gDoc, currentPage, bookmark.name);
            if (annot) {
                results.success++;
                console.println("✓ 成功添加页脚注释");
            }
        } catch(error) {
            results.fail++;
            var errorInfo = {
                index: i + 1,
                name: bookmark.name,
                error: error.toString()
            };
            results.failedBookmarks.push(errorInfo);
            
            // 记录错误日志
            errorLog += "\n书签处理失败 #" + (i + 1) + ":\n";
            errorLog += "书签名称: " + bookmark.name + "\n";
            errorLog += "错误信息: " + error.toString() + "\n";
            console.println("✗ 处理失败: " + error.toString());
        }
    }
    
    // 列出没有书签的页面
    var unprocesedPages = [];
    for (var j = 0; j < processedPages.length; j++) {
        if (!processedPages[j]) {
            unprocesedPages.push(j + 1);
        }
    }
    
    console.println("\n=== 处理完成 ===");
    console.println(`总页数: ${gDoc.numPages}`);
    console.println(`处理成功: ${results.success}`);
    console.println(`处理失败: ${results.fail}`);
    
    if (unprocesedPages.length > 0) {
        console.println("\n=== 未处理页面 ===");
        console.println(`共 ${unprocesedPages.length} 页没有对应书签:`);
        console.println(unprocesedPages.join(", "));
    }
    
    // 修改最终提示对话框
    if (results.failedBookmarks.length > 0) {
        console.println("\n=== 失败详情 ===");
        console.println(errorLog);
        
        var failureMessage = "处理失败的书签:\n\n";
        for (var j = 0; j < results.failedBookmarks.length; j++) {
            var failed = results.failedBookmarks[j];
            failureMessage += "序号 " + failed.index + ":\n名称: " + failed.name + "\n错误: " + failed.error + "\n\n";
        }
        
        app.alert({
            cMsg: "处理完成！\n成功: " + results.success + " 个\n失败: " + results.fail + " 个\n\n" + failureMessage,
            cTitle: "添加页脚注释 - 包含失败信息",
            nIcon: 2,
            nType: 0
        });
    } else {
        app.alert({
            cMsg: `处理完成！\n\n` +
                  `总页数: ${gDoc.numPages}\n` +
                  `成功处理: ${results.success} 个书签\n` +
                  `未处理页面: ${unprocesedPages.length} 页\n\n` +
                  (unprocesedPages.length > 0 ? 
                   `未处理的页码: ${unprocesedPages.join(", ")}` : 
                   "所有页面都已处理！"),
            cTitle: "添加页脚注释",
            nIcon: 3,
            nType: 0
        });
    }
}

/**
 * 添加书签页脚
 * 主要流程：
 * 1. 获取当前文档
 * 2. 检查并获取书签
 * 3. 创建预览
 * 4. 等待用户确认
 */
function addBookmarkFooters() {
    try {
        console.println("=== 开始执行脚本 ===");
        console.println("时间: " + new Date().toLocaleString());
        
        gDoc = app.activeDocs[0];
        if (!gDoc) {
            throw new Error("未找到活动文档");
        }
        console.println("文档名称: " + gDoc.documentFileName);
        
        var root = gDoc.bookmarkRoot;
        gBookmarks = root.children;
        if (!gBookmarks || gBookmarks.length === 0) {
            throw new Error("文档中未找到书签");
        }
        console.println("检测到书签数量: " + gBookmarks.length);
        
        gOriginalPage = gDoc.pageNum;
        var previewBookmark = gBookmarks[0];
        previewBookmark.execute();
        
        // 先显示提示，让用户知道预览将要添加
        app.alert({
            cMsg: "即将在第一个书签页面添加预览页脚\n\n" +
                  "位置：距底部 " + CONFIG.annotation.rect[1] + "-" + CONFIG.annotation.rect[3] + " 点\n" +
                  "字体：" + CONFIG.annotation.font + "，" + CONFIG.annotation.textSize + "号\n" +
                  "对齐：居中\n" +
                  "示例文本：" + previewBookmark.name,
            cTitle: "页脚注释预览准备",
            nIcon: 3,
            nType: 0
        });
        
        // 添加预览注释
        gPreviewAnnot = createAnnotation(gDoc, gDoc.pageNum, previewBookmark.name);
        
        // 延迟显示确认对话框
        app.setTimeOut("showConfirmDialog()", 5000);
        
    } catch(e) {
        console.println("\n=== 发生错误 ===");
        console.println("错误详情: " + e);
        app.alert("发生错误: " + e);
    }
}

/**
 * 批量处理所有书签
 * @param {Array} bookmarks - 书签数组
 * @returns {Promise} 处理完成的Promise
 */
async function processAllBookmarks(bookmarks) {
    console.log('\n=== 开始处理所有书签 ===');
    console.log(`时间: ${new Date().toLocaleString()}`);
    console.log(`总书签数量: ${bookmarks.length}\n`);
    
    const startTime = Date.now();
    let successCount = 0;
    let failCount = 0;

    for (let i = 0; i < bookmarks.length; i++) {
        console.log(`\n--- 处理书签 [${i + 1}/${bookmarks.length}] ---`);
        try {
            await processBookmark(bookmarks[i]);
            successCount++;
        } catch (error) {
            console.error(`处理失败: ${error.message}`);
            failCount++;
        }
    }

    const endTime = Date.now();
    const duration = ((endTime - startTime) / 1000).toFixed(2);
    
    console.log('\n=== 处理完成 ===');
    console.log(`总耗时: ${duration}秒`);
    console.log(`成功数量: ${successCount}`);
    console.log(`失败数量: ${failCount}`);
}

// 启动脚本
addBookmarkFooters();