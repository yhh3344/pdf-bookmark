/**
 * PDF注释清理工具
 * 用于清理PDF文档中的各类注释，包括：
 * - 水印注释
 * - 打字机注释
 * - 页眉页脚
 * - 其他所有类型的注释
 * 
 * @copyright Copyright (c) 2025 yhh
 * @license MIT
 * @version 1.0.0
 * @tested Foxit PDF Editor 13.0.1.21693
 */

// 全局变量定义
var gDoc;          // 当前活动的PDF文档

/**
 * 清理指定页面的所有注释
 * @param {Object} doc - PDF文档对象
 * @param {number} pageNum - 页码
 * @returns {number} 清理的注释数量
 */
function cleanPageAnnotations(doc, pageNum) {
    try {
        var annots = doc.getAnnots(pageNum);
        if (!annots || annots.length === 0) {
            return 0;
        }

        var count = 0;
        for (var i = annots.length - 1; i >= 0; i--) {
            var annot = annots[i];
            annot.destroy();
            count++;
        }
        return count;
    } catch (e) {
        console.println(`清理第 ${pageNum + 1} 页注释时出错: ${e}`);
        throw e;
    }
}

/**
 * 清理所有页面的注释
 * 主要流程：
 * 1. 确认用户操作
 * 2. 遍历所有页面
 * 3. 清理每页注释
 * 4. 显示清理报告
 */
function cleanAllAnnotations() {
    console.println("=== 开始执行脚本 ===");
    console.println("时间: " + new Date().toLocaleString());
    
    try {
        // 获取当前文档
        gDoc = app.activeDocs[0];
        if (!gDoc) {
            throw new Error("未找到活动文档");
        }
        console.println("文档名称: " + gDoc.documentFileName);
        
        // 确认操作
        var response = app.alert({
            cMsg: "此操作将清除文档中的所有注释，包括：\n" +
                  "- 水印\n" +
                  "- 打字机注释\n" +
                  "- 页眉页脚\n" +
                  "- 其他所有类型的注释\n\n" +
                  "此操作不可撤销，是否继续？",
            cTitle: "确认清理注释",
            nIcon: 2,
            nType: 2
        });
        
        if (response !== 4) {
            console.println("用户取消操作");
            return;
        }
        
        // 开始清理
        console.println("\n=== 开始清理注释 ===");
        var totalPages = gDoc.numPages;
        var totalAnnots = 0;
        var failedPages = [];
        
        for (var i = 0; i < totalPages; i++) {
            try {
                console.println(`\n处理第 ${i + 1}/${totalPages} 页...`);
                var count = cleanPageAnnotations(gDoc, i);
                totalAnnots += count;
                console.println(`✓ 已清理 ${count} 个注释`);
            } catch (error) {
                failedPages.push(i + 1);
                console.println(`✗ 处理失败: ${error}`);
            }
        }
        
        // 显示结果
        console.println("\n=== 清理完成 ===");
        console.println(`总页数: ${totalPages}`);
        console.println(`清理注释总数: ${totalAnnots}`);
        
        var resultMessage = `清理完成！\n\n` +
                           `总页数: ${totalPages}\n` +
                           `清理注释总数: ${totalAnnots}\n`;
        
        if (failedPages.length > 0) {
            resultMessage += `\n处理失败页数: ${failedPages.length}\n` +
                           `失败页码: ${failedPages.join(", ")}\n`;
        }
        
        app.alert({
            cMsg: resultMessage,
            cTitle: "注释清理结果",
            nIcon: failedPages.length > 0 ? 2 : 3,
            nType: 0
        });
        
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

// 启动脚本
cleanAllAnnotations(); 