# !/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
# Func       ：数据填充工具（Streamlit 版）
"""
import streamlit as st
from main import *

class excel2wordSt:

    def run(self):
        # ---------------- 页面配置（必须是第一个 st 调用） ----------------
        st.set_page_config(
            page_title="数据填充工具",
            page_icon="📄",
            layout="wide",
        )

        # ---------------- 会话状态初始化 ----------------
        if "result" not in st.session_state:
            st.session_state.result = None

        # ---------------- 标题 ----------------
        st.markdown(
            "<h1 style='text-align:center;color:#2c7be5;'>📄 数据填充工具</h1>",
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # ---------------- 主体布局（中间一列） ----------------
        left, center, right = st.columns([1, 3, 1])

        with center:
            with st.container(border=True):
                # ---------- 1. 选择数据源 Excel ----------
                st.markdown("**📊 选择数据源（Excel）**")
                excel_file = st.file_uploader(
                    "上传 Excel",
                    type=["xlsx"],
                    key="excel_uploader",
                    label_visibility="collapsed",
                )
                if excel_file is not None:
                    st.success(f"✅ 已选择：{excel_file.name}")

                st.divider()

                # ---------- 2. 选择模板 Word ----------
                st.markdown("**📝 选择模板（Word）**")
                word_file = st.file_uploader(
                    "上传 Word",
                    type=["docx"],
                    key="word_uploader",
                    label_visibility="collapsed",
                )
                if word_file is not None:
                    st.success(f"✅ 已选择：{word_file.name}")

                st.divider()

                # ---------- 3. 指定输出位置 ----------
                st.markdown("**📁 输出文件夹路径**")
                output_path = st.text_input(
                    "输出文件夹路径",
                    value="",
                    placeholder="例如：C:/output 或 ./output",
                    label_visibility="collapsed",
                )
                st.caption(
                    "提示：Streamlit 运行在服务端，无法像桌面程序那样弹出本地文件夹选择框；"
                    "请直接填写路径（支持相对路径，相对于程序运行目录）。留空则输出到当前目录。"
                )

                st.divider()

                # ---------- 4. 开始运行按钮 ----------
                ready = (excel_file is not None) and (word_file is not None)
                run_clicked = st.button(
                    "🚀 开始运行",
                    type="primary",
                    use_container_width=True,
                    disabled=not ready,
                )

            # ---------------- 5. 处理逻辑 ----------------
            if run_clicked and ready:
                base_dir = os.getcwd()
                raw_path = (output_path or "").strip()
                if raw_path:
                    output_dir = raw_path if os.path.isabs(raw_path) else os.path.join(base_dir, raw_path)
                else:
                    output_dir = base_dir

                # 确保目录存在
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except Exception as e:
                    st.session_state.result = {"error": f"创建输出目录失败：{e}"}
                else:
                    # 生成输出文件名
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    output_filename = f"填充结果_{timestamp}.docx"
                    output_full_path = os.path.join(output_dir, output_filename)

                    # 模拟处理：实际可在这里调用 python-docx / openpyxl 做真正的填充
                    content = (
                        f"模拟填充结果\n"
                        f"数据源：{excel_file.name}\n"
                        f"模板：{word_file.name}\n"
                        f"时间：{timestamp}\n"
                    ).encode("utf-8")

                    try:
                        with open(output_full_path, "wb") as f:
                            f.write(content)
                    except Exception as e:
                        st.session_state.result = {"error": f"文件生成失败：{e}"}
                    else:
                        st.session_state.result = {
                            "error": None,
                            "excel_name": excel_file.name,
                            "word_name": word_file.name,
                            "output_dir": output_dir,
                            "output_filename": output_filename,
                            "output_full_path": output_full_path,
                            "content": content,
                        }

            # ---------------- 结果展示 ----------------
            result = st.session_state.result
            if result:
                if result.get("error"):
                    st.error(f"❌ {result['error']}")
                else:
                    with st.container(border=True):
                        st.success("✅ 处理完成！")
                        st.markdown(f"📊 **数据源文件：** {result['excel_name']}")
                        st.markdown(f"📝 **模板文件：** {result['word_name']}")
                        st.markdown(f"📁 **输出文件夹：** {result['output_dir']}")
                        st.markdown(f"📄 **生成文件：** {result['output_filename']}")
                        st.markdown(f"📂 **完整路径：** {result['output_full_path']}")

                        # 额外提供下载（模拟内容，实际可按真实文件字节流下载）
                        st.download_button(
                            "⬇️ 下载生成文件",
                            data=result["content"],
                            file_name=result["output_filename"],
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )
