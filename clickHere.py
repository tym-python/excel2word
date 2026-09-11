#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
# File       : app.py
# Func       ：数据填充工具（Streamlit 界面）
"""
import os,sys
import time
import subprocess
import streamlit as st

class excel2wordSt:

    # ================= 会话状态初始化 =================
    def _init_state(self):
        defaults = {
            "result": None,  # 处理结果
            "running": False,  # 是否正在处理
            "last_run_time": 0.0,  # 上次运行时间（用于双击冷却）
            "uploader_round": 0,      # ★ file_uploader 的"轮次"，用于重置控件
        }
        for k, v in defaults.items():
            st.session_state.setdefault(k, v)

    # ================= 辅助函数 =================
    def open_folder_in_os(self, path: str) -> bool:
        """
        尝试在操作系统层面打开文件夹。
        仅对「本地部署」有意义——如果 Streamlit 跑在远程服务器上，
        这里打开的是服务器上的文件夹，用户看不到。
        """
        if not path or not os.path.isdir(path):
            return False
        try:
            if sys.platform == "win32":
                os.startfile(path)  # noqa: 仅在 Windows 上存在
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
            return True
        except Exception:
            return False

    def reset_all(self):
        """清空所有会话状态，相当于重新打开页面。"""
        st.session_state.result = None
        st.session_state.running = False
        st.session_state.last_run_time = 0.0
        st.session_state.uploader_round += 1  # ★ 换 key → 上传控件被重建

    def run(self):
        # ---------------- 页面配置（必须是第一个 st 调用） ----------------
        st.set_page_config(
            page_title="数据填充工具",
            page_icon="📄",
            layout="wide",
        )

        # ---------------- 会话状态初始化 ----------------
        self._init_state()
        COOLDOWN_SECONDS = 1.5  # 双击冷却时间

        # ---------------- 标题 ----------------
        st.markdown(
            "<h1 style='text-align:center;color:#2c7be5;'>📄 数据填充工具</h1>",
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # ---------------- 主体布局 ----------------
        left, center, right = st.columns([1, 3, 1])
        round_ = st.session_state.uploader_round

        with center:
            with st.container(border=True):
                # ---------- 1. 选择数据源 Excel ----------
                st.markdown("**📊 选择数据源（Excel）**")
                excel_file = st.file_uploader(
                    "上传 Excel",
                    type=["xlsx"],
                    key=f"excel_uploader_{round_}",
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
                    key=f"word_uploader_{round_}",
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
                    placeholder="留空 → 默认 app.py 所在目录下的 output 文件夹",
                    label_visibility="collapsed",
                )
                st.caption(
                    "留空则输出到「运行入口文件所在目录 / output」。"
                    "也可以填写绝对路径（如 D:/result）或相对于入口目录的相对路径（如 result）。"
                )

                st.divider()

                # ---------- 4. 开始运行按钮 ----------
                ready = (excel_file is not None) and (word_file is not None)

                # 冷却判断：上一次点击在 COOLDOWN_SECONDS 内则按钮不可点
                in_cooldown = (time.time() - st.session_state.last_run_time) < COOLDOWN_SECONDS
                button_disabled = (not ready) or st.session_state.running or in_cooldown

                if st.session_state.running:
                    label = "⏳ 正在处理..."
                elif in_cooldown:
                    label = "⏳ 请稍候..."
                else:
                    label = "🚀 开始运行"

                run_clicked = st.button(
                    label,
                    type="primary",
                    use_container_width=True,
                    disabled=button_disabled,
                    key="run_btn",
                )

                # ---------- 5. 清空重来 ----------
                if st.button(
                        "🔄 清空重来",
                        use_container_width=True,
                        disabled=st.session_state.running,
                ):
                    self.reset_all()
                    st.rerun()


            # ---------------- 6. 处理逻辑 ----------------
            if run_clicked and ready and not st.session_state.running:
                st.session_state.running = True
                st.session_state.last_run_time = time.time()
                st.session_state.result = None

                try:
                    with st.spinner("正在填充数据..."):
                        from excel2word import excel2word

                        # ---- 决定输入目录与输出目录 ----
                        # 输入目录：用「入口文件所在目录」（默认策略，无需用户操作）
                        # 输出目录：用户留空 → 入口目录/output；填写 → 用户指定
                        c_probe = excel2word(input_DIR_NAME='')     # 仅用于探测 base_dir
                        base_dir = c_probe.base_dir

                        raw_path = (output_path or "").strip()
                        if raw_path:
                            if os.path.isabs(raw_path):
                                output_dir = os.path.abspath(raw_path)
                            else:
                                output_dir = os.path.abspath(os.path.join(base_dir, raw_path))
                        else:
                            output_dir = os.path.join(base_dir, "output")
                        os.makedirs(output_dir, exist_ok=True)

                        # ---- 上传文件落盘到 base_dir，供 excel2word 读取 ----
                        with open(os.path.join(base_dir, excel_file.name), "wb") as f:
                            f.write(excel_file.getbuffer())
                        with open(os.path.join(base_dir, word_file.name), "wb") as f:
                            f.write(word_file.getbuffer())

                        # ---- 调用 excel2word ----
                        # input_DIR_NAME=''          → 输入 = base_dir（上传文件落地处）
                        # OUTPUT_DIR_NAME=output_dir → 输出 = 用户指定目录
                        # Image_DIR_NAME             → 跟随输出目录下的 img
                        c = excel2word(
                            input_DIR_NAME='',
                            OUTPUT_DIR_NAME=output_dir,
                            Image_DIR_NAME=os.path.join(base_dir, "img"),
                        )
                        generated = c.main(excelName=excel_file.name, wordName=word_file.name)

                        # ---- 收集输出目录内容 ----
                        if generated < 0:
                            st.session_state.result = {
                                "error": "excel2word 执行失败。"
                            }
                        else:
                            st.session_state.result = {
                                "error": None,
                                "output_dir": output_dir,
                                "excel_name": excel_file.name,
                                "word_name": word_file.name,
                                "generated": generated,
                            }

                except ImportError:
                    st.session_state.result = {
                        "error": "未找到 excel2word 模块，请确认 excel2word.py 与 app.py 位于同一目录。"
                    }
                except Exception as e:
                    st.session_state.result = {"error": f"处理失败：{e}"}
                finally:
                    st.session_state.running = False

            # ---------------- 结果展示 ----------------
            result = st.session_state.result
            if result:
                if result.get("error"):
                    st.error(f"❌ {result['error']}")
                    if st.button("🔄 清空重来", use_container_width=True, key="reset_after_error"):
                        self.reset_all()
                        st.rerun()
                else:
                    with st.container(border=True):
                        st.success(f"✅ 处理完成！共生成 {result.get('generated', 0)} 个文档。")
                        st.markdown(f"📊 **数据源文件：** {result['excel_name']}")
                        st.markdown(f"📝 **模板文件：** {result['word_name']}")
                        st.markdown(f"📁 **输出文件夹：** {result['output_dir']}")
                        if st.button("📂 打开输出文件夹", use_container_width=True, key="open_folder_btn"):
                            if self.open_folder_in_os(result["output_dir"]):
                                st.toast("已在服务端打开文件夹", icon="📂")
                            else:
                                st.warning(
                                    "无法在当前运行环境中打开文件夹（可能是远程服务器），"
                                    "请复制上方路径到文件管理器中手动打开。"
                                )
                        # ---- 处理下一批 ----
                        if st.button("🔄 处理下一批", use_container_width=True, key="next_batch_btn"):
                            self.reset_all()
                            st.rerun()

if __name__ == '__main__':
    app = excel2wordSt()
    app.run()
# streamlit run clickHere.py