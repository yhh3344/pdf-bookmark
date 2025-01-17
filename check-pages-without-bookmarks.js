// 全局变量
var gDoc;
var gBookmarks;

// 检查无书签页面的函数
function findPagesWithoutBookmarks(doc, bookmarks) {
    console.println("\n=== 开始检查无书签页面 ===");
    var pagesWithoutBookmarks = [];
    var bookmarkPages = {};
    var originalPage = doc.pageNum;
    var duplicatePages = {};
    
    try {
        // 记录所有书签对应的页码
        console.println("\n=== 书签页码统计 ===");
        console.println("总页数：" + doc.numPages);
        console.println("书签数：" + bookmarks.length);
        console.println("\n正在收集书签页码...");
        
        for (var i = 0; i < bookmarks.length; i++) {
            var bookmark = bookmarks[i];
            bookmark.execute();
            var pageNum = doc.pageNum;
            
            // 检查重复书签页
            if (bookmarkPages[pageNum]) {
                if (!duplicatePages[pageNum]) {
                    duplicatePages[pageNum] = [];
                }
                duplicatePages[pageNum].push(bookmark.name);
            }
            
            bookmarkPages[pageNum] = bookmark.name;
            console.println("书签 [" + (i + 1) + "/" + bookmarks.length + "]: " + bookmark.name + " -> 页码 " + (pageNum + 1));
        }
        
        // 检查重复书签页
        var hasDuplicates = Object.keys(duplicatePages).length > 0;
        if (hasDuplicates) {
            console.println("\n=== 发现重复书签页 ===");
            for (var page in duplicatePages) {
                console.println("第 " + (parseInt(page) + 1) + " 页有多个书签：");
                console.println("- " + bookmarkPages[page]);
                duplicatePages[page].forEach(function(name) {
                    console.println("- " + name);
                });
            }
        }
        
        // 检查每一页是否有书签
        console.println("\n=== 无书签页面检查 ===");
        for (var pageNum = 0; pageNum < doc.numPages; pageNum++) {
            if (!bookmarkPages[pageNum]) {
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
            resultMessage += "发现重复书签页！请查看控制台获取详细信息\n\n";
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

// 主函数
function checkPagesWithoutBookmarks() {
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
        
        // 执行检查
        findPagesWithoutBookmarks(gDoc, gBookmarks);
        
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
checkPagesWithoutBookmarks(); 