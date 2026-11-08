"""个性化功能定制分区。"""

from __future__ import annotations

from app.views.placeholder import PlaceholderPage


class CustomizationView(PlaceholderPage):
    def __init__(self, parent=None) -> None:
        super().__init__(
            title="个性化功能定制",
            description="按岗位习惯定制工作台入口、看板字段与消息提醒方式。"
            "该分区将在后续版本开放。",
            features=(
                "自定义首页快捷入口与卡片布局",
                "计划看板展示字段与排序列配置",
                "同步提醒与异常消息偏好设置",
            ),
            icon_kind="custom",
            parent=parent,
        )
