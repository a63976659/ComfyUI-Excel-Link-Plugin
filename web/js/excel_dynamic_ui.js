import { app } from "../../../scripts/app.js";

// 注册扩展以实现前端动态交互
app.registerExtension({
    name: "ComfyUI.ExcelLink.DynamicUI",
    async nodeCreated(node) {
        // 匹配后端节点类名
        if (node.comfyClass === "图片插入表格") {
            
            // 内部辅助函数：根据名称查找组件
            const getWidget = (name) => node.widgets.find(w => w.name === name);
            
            const modeWidget = getWidget("缩放模式");
            const widthWidget = getWidget("图片宽度");
            const heightWidget = getWidget("图片高度");
            const ratioWidget = getWidget("缩放比例");

            // 定义刷新显示状态的逻辑
            const refreshWidgets = () => {
                if (!modeWidget) return;
                
                const mode = modeWidget.value;

                // 1. 仅在“固定尺寸”模式下显示宽度和高度
                const isFixed = mode === "固定尺寸";
                if (widthWidget) widthWidget.type = isFixed ? "INT" : "converted-widget";
                if (heightWidget) heightWidget.type = isFixed ? "INT" : "converted-widget";

                // 2. 仅在“按比例缩放”模式下显示缩放比例
                const isRatio = mode === "按比例缩放";
                if (ratioWidget) ratioWidget.type = isRatio ? "FLOAT" : "converted-widget";

                // 3. 在“匹配单元格”或“原图大小”模式下，上述参数都会被隐藏

                // 触发画布重绘以更新 UI 布局
                app.canvas.draw(true, true);
            };

            // 监听模式切换事件
            modeWidget.callback = () => {
                refreshWidgets();
            };

            // 节点创建后延迟执行一次，确保初始状态正确
            setTimeout(refreshWidgets, 100);
        }
    }
});