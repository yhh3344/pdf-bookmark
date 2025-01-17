// 全局变量
var gDoc;
var gBookmarks;

// 检查无书签页面的函数
function findPagesWithoutBookmarks(doc, bookmarks) {
    console.println("\n=== 开始检查无书签页面 ===");
    var pagesWithoutBookmarks = [];
    var bookmarkPages = {};
    var originalPage = doc.pageNum;
    
    try {
        // 记录所有书签对应的页码
        console.println("正在收集书签页码...");
        for (var i = 0; i < bookmarks.length; i++) {
            var bookmark = bookmarks[i];
            bookmark.execute();
            bookmarkPages[doc.pageNum] = true;
            console.println("书签 [" + (i + 1) + "/" + bookmarks.length + "]: " + bookmark.name + " -> 页码 " + (doc.pageNum + 1));
        }
        
        // 检查每一页是否有书签
        console.println("\n正在检查无书签页面...");
        for (var pageNum = 0; pageNum < doc.numPages; pageNum++) {
            if (!bookmarkPages[pageNum]) {
                pagesWithoutBookmarks.push(pageNum + 1);
            }
        }
        
        // 显示结果
        if (pagesWithoutBookmarks.length > 0) {
            console.println("\n发现 " + pagesWithoutBookmarks.length + " 个无书签页面:");
            console.println("页码: " + pagesWithoutBookmarks.join(", "));
            
            app.alert({
                cMsg: "发现 " + pagesWithoutBookmarks.length + " 个无书签页面：\n\n" +
                      "页码：" + pagesWithoutBookmarks.join(", ") + "\n\n" +
                      "详细信息已输出到控制台",
                cTitle: "无书签页面检查结果",
                nIcon: 2,
                nType: 0
            });
        } else {
            console.println("\n未发现无书签页面");
            app.alert({
                cMsg: "文档中所有页面都有对应的书签！",
                cTitle: "无书签页面检查结果",
                nIcon: 3,
                nType: 0
            });
        }
        
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