/**
 * PDF书签分析工具
 * 用于分析PDF文档中的书签分布情况，包括：
 * - 检测无书签页面
 * - 统计重复书签页
 * - 生成书签分布报告
 * 
 * @copyright Copyright (c) 2025 yhh
 * @license MIT
 * @version 1.0.0
 * @tested Foxit PDF Editor 13.0.1.21693
 */

// 全局变量定义
var gDoc;          // 当前活动的PDF文档
var gBookmarks;    // 文档中的所有书签

/**
 * 分析书签分布情况
 * @param {Object} doc - PDF文档对象
 * @param {Array} bookmarks - 书签数组
 */
function analyzeBookmarkDistribution(doc, bookmarks) {
    console.println("\n=== 开始分析书签分布 ===");
    var pagesWithoutBookmarks = [];
    var bookmarkPages = {};
    var pagesWithBookmarks = {};  // 用于检查页面是否有书签
    var originalPage = doc.pageNum;
    
    try {
        // 修改数据结构，使每个页码可以存储多个书签
        console.println("\n=== 书签页码统计 ===");
        console.println("总页数：" + doc.numPages);
        console.println("书签数：" + bookmarks.length);
        console.println("\n正在收集书签页码...");
        
        for (var i = 0; i < bookmarks.length; i++) {
            var bookmark = bookmarks[i];
            bookmark.execute();
            var pageNum = doc.pageNum;
            
            // 初始化数组（如果不存在）
            if (!bookmarkPages[pageNum]) {
                bookmarkPages[pageNum] = [];
            }
            
            // 添加书签到对应页码
            bookmarkPages[pageNum].push(bookmark.name);
            pagesWithBookmarks[pageNum] = true;
            
            console.println("书签 [" + (i + 1) + "/" + bookmarks.length + "]: " + bookmark.name + " -> 页码 " + (pageNum + 1));
        }
        
        // 检查重复书签页
        var duplicatePages = {};
        var hasDuplicates = false;
        
        for (var page in bookmarkPages) {
            if (bookmarkPages[page].length > 1) {
                hasDuplicates = true;
                duplicatePages[page] = bookmarkPages[page];
            }
        }
        
        if (hasDuplicates) {
            console.println("\n=== 发现重复书签页 ===");
            for (var page in duplicatePages) {
                console.println("第 " + (parseInt(page) + 1) + " 页有 " + duplicatePages[page].length + " 个书签：");
                duplicatePages[page].forEach(function(name) {
                    console.println("- " + name);
                });
            }
        }
        
        // 检查无书签页面
        console.println("\n=== 无书签页面检查 ===");
        for (var pageNum = 0; pageNum < doc.numPages; pageNum++) {
            if (!pagesWithBookmarks[pageNum]) {
                pagesWithoutBookmarks.push(pageNum + 1);
            }
        }
        
        // 显示结果
        console.println("\n=== 检查结果汇总 ===");
        console.println("总页数：" + doc.numPages);
        console.println("书签数：" + bookmarks.length);
        console.println("无书签页面数：" + pagesWithoutBookmarks.length);
        console.println("重复书签页数：" + Object.keys(duplicatePages).length);
        
        var resultMessage = "检查结果：\n\n" +
                           "总页数：" + doc.numPages + "\n" +
                           "书签数：" + bookmarks.length + "\n" +
                           "无书签页面数：" + pagesWithoutBookmarks.length + "\n" +
                           "重复书签页数：" + Object.keys(duplicatePages).length + "\n\n";
        
        if (pagesWithoutBookmarks.length > 0) {
            console.println("\n无书签页面：" + pagesWithoutBookmarks.join(", "));
            resultMessage += "无书签页面：" + pagesWithoutBookmarks.join(", ") + "\n\n";
        }
        
        if (hasDuplicates) {
            var duplicatePagesMsg = "重复书签页：\n";
            for (var page in duplicatePages) {
                duplicatePagesMsg += "第 " + (parseInt(page) + 1) + " 页 (" + 
                                   duplicatePages[page].length + " 个书签):\n";
                duplicatePages[page].forEach(function(name) {
                    duplicatePagesMsg += "- " + name + "\n";
                });
                duplicatePagesMsg += "\n";
            }
            console.println("\n" + duplicatePagesMsg);
            resultMessage += duplicatePagesMsg;
        }
        
        app.alert({
            cMsg: resultMessage + "详细信息已输出到控制台",
            cTitle: "书签检查结果",
            nIcon: 2,
            nType: 0
        });
        
    } catch(e) {
        console.println("检查过程中发生错误: " + e);
        app.alert({
            cMsg: "检查过程中发生错误：\n" + e,
            cTitle: "错误",
            nIcon: 0,
            nType: 0
        });
    } finally {
        // 恢复原始页面
        doc.pageNum = originalPage;
    }
}

/**
 * 启动书签分析
 * 主要流程：
 * 1. 获取当前活动文档
 * 2. 检查并获取书签
 * 3. 执行分析
 */
function startBookmarkAnalysis() {
    console.println("=== 开始执行脚本 ===");
    console.println("时间: " + new Date().toLocaleString());
    
    try {
        // 获取当前文档
        gDoc = app.activeDocs[0];
        if (!gDoc) {
            throw new Error("未找到活动文档");
        }
        console.println("文档名称: " + gDoc.documentFileName);
        
        // 获取书签
        var root = gDoc.bookmarkRoot;
        gBookmarks = root.children;
        if (!gBookmarks || gBookmarks.length === 0) {
            throw new Error("文档中未找到书签");
        }
        console.println("检测到书签数量: " + gBookmarks.length);
        
        // 执行分析
        analyzeBookmarkDistribution(gDoc, gBookmarks);
        
    } catch(e) {
        console.println("\n=== 发生错误 ===");
        console.println("错误详情: " + e);
        app.alert({
            cMsg: "发生错误：\n" + e,
            cTitle: "错误",
            nIcon: 0,
            nType: 0
        });
    }
}

// 执行脚本
startBookmarkAnalysis(); 