import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import json
from openai import OpenAI
from datetime import datetime

# 1. 网页基础配置
st.set_page_config(page_title="体卫艺办公助手", page_icon="📋", layout="wide", initial_sidebar_state="collapsed")

# --- 🎨 深度美化 / CSS 设计 ---
st.markdown("""
<style>
    /* --- 全局基础设定 --- */
    html, body, [class*="css"] {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif, 'Apple Color Emoji', 'Segoe UI Emoji', 'Segoe UI Symbol';
        -webkit-font-smoothing: antialiased;
        -webkit-text-size-adjust: 100%; /* 禁止 iOS 自动调整字体大小 */
    }
    
    /* 页面背景 */
    .stApp {
        background-color: #f7f9fc;
        background-image: linear-gradient(180deg, #f7f9fc 0%, #edf2f7 100%);
    }

    /* --- 核心优化：将 Radio 按钮变成了“高级导航卡片” --- */
    /* 1. 隐藏“功能切换”这个Label，太占地方 */
    div[data-testid="stRadio"] > label {
        display: none !important;
    }

    /* 2. 容器设置：更紧凑 */
    div[data-testid="stRadio"] > div {
        gap: 8px !important; /* 间距由 12px 缩减到 8px */
        margin-top: 0px !important;
    }
    
    /* 3. 卡片基础样式 - 去掉边框，纯净风格 */
    div[data-testid="stRadio"] label {
        background-color: #f8f9fa !important; /* 默认淡灰背景，不显眼 */
        padding: 12px 16px !important; /* 内边距稍微调小，更精致 */
        border-radius: 10px !important;
        border: 1px solid transparent !important; /* 默认无边框 */
        transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
        margin-bottom: 0px !important;
        color: #555 !important;
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
        cursor: pointer;
        position: relative; /* 为了隐藏圆圈做定位 */
    }

    /* 4. 彻底隐藏原生的圆圈和小圆点 - 使用更猛的 CSS hack */
    div[data-testid="stRadio"] div[role="radiogroup"] > label > div:first-child {
        display: none !important;
    }
    
    /* 5. 选中状态 - 真正的“高级感”：深蓝底白字 */
    /* 由于 data-checked 可能不稳定，这里使用 :has() 选择器 (现代浏览器支持) */
    div[data-testid="stRadio"] label:has(input:checked) {
        background: linear-gradient(135deg, #1565c0 0%, #0d47a1 100%) !important;
        color: #ffffff !important;
        font-weight: 500 !important;
        box-shadow: 0 4px 12px rgba(21, 101, 192, 0.3) !important;
        transform: translateY(-1px);
        border: none !important;
    }
    
    /* 选中状态文字 */
    div[data-testid="stRadio"] label:has(input:checked) p {
        color: #ffffff !important;
    }
    
    /* 降级兼容：如果 :has 不支持，尝试使用 sibling 选择器 */
    div[data-testid="stRadio"] input:checked + div {
        background: linear-gradient(135deg, #1565c0 0%, #0d47a1 100%) !important;
        color: white !important;
    }


    /* --- 侧边栏优化 --- */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e0e0e0;
    }
    
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #1e88e5 !important;
    }

    /* --- 顶部 Header --- */
    header[data-testid="stHeader"] {
        background: rgba(255, 255, 255, 0.95) !important;
        backdrop-filter: blur(10px);
        border-bottom: 1px solid #f0f0f0;
        height: 60px !important;
        z-index: 999999;
    }

    /* --- 内容容器美化 --- */
    div[data-testid="stVerticalBlockBorderWrapper"] > div {
        background-color: #ffffff;
        border: 1px solid #e8ebf0 !important;
        border-radius: 16px !important; 
        box-shadow: 0 4px 20px rgba(0,0,0,0.04);
        padding: 1.5rem !important;
    }

    /* --- 输入控件优化 --- */
    .stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"] {
        border-radius: 10px !important;
        border: 1px solid #cfd8dc !important;
        min-height: 50px !important;
        font-size: 16px !important;
        line-height: 1.5 !important;
        padding-left: 12px !important;
    }
    
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: #2196f3 !important;
        box-shadow: 0 0 0 2px rgba(33, 150, 243, 0.2) !important;
    }

    /* --- 按钮美化 --- */
    div.stButton > button {
        border-radius: 24px !important;
        font-weight: 600 !important;
        min-height: 48px !important;
        font-size: 16px !important;
        padding: 0 24px !important;
        width: 100%;
    }
    
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #1976d2 0%, #1e88e5 100%) !important;
        box-shadow: 0 6px 16px rgba(25, 118, 210, 0.3) !important;
        border: none !important;
    }

    /* --- 核心优化：针对移动设备 (iPhone 17 Pro Max / WeChat) --- */
    @media (max-width: 768px) {
        /* 1. 强制主内容区全宽，增加顶部安全间距，防止微信标题栏遮挡 */
        .block-container {
            padding-top: 5rem !important; 
            padding-left: 0.8rem !important;
            padding-right: 0.8rem !important;
            padding-bottom: 5rem !important;
            max-width: 100vw !important;
        }

        /* 2. 侧边栏优化：在超大屏 iPhone 上稍微变宽一点，更易操作 */
        section[data-testid="stSidebar"] {
            width: 75vw !important;
            min-width: 280px !important;
        }

        /* 3. 增大所有交互按钮的点击面积 (Fat Finger Friendly) */
        div.stButton > button {
            height: 56px !important;
            font-size: 18px !important;
        }

        /* 4. Radio 卡片自适应：在小屏上堆叠，在大屏上并排 */
        div[data-testid="stRadio"] label {
            padding: 16px !important;
            font-size: 16px !important;
            margin-bottom: 4px !important;
        }
        
        /* 5. 隐藏 Streamlit 自带的底部装饰条，留出更多可视空间 */
        footer { visibility: hidden; }
        #MainMenu { visibility: hidden; }
    }
</style>
""", unsafe_allow_html=True)

# --- 🔒 通讯录专属密码 ---
CONTACT_PASSWORD = "lhjy" 

# 2. 核心配置
MY_API_KEY = "sk-dzsawqzsktjximglmkzyezbtyhqbysvenoxublemcgertlqp"
BASE_URL = "https://api.siliconflow.cn/v1"

# 初始化状态
if "contacts_authenticated" not in st.session_state:
    st.session_state.contacts_authenticated = False
if "parseddata_doc" not in st.session_state:
    st.session_state.parseddata_doc = None
# 新增：两步流程的状态管理
if "step" not in st.session_state:
    st.session_state.step = 1  # 1=输入, 2=确认润色, 3=确认字段
if "polished_text" not in st.session_state:
    st.session_state.polished_text = None
if "original_input" not in st.session_state:
    st.session_state.original_input = ""

# 3. 侧边栏导航
with st.sidebar:
    st.header("⚙️ 体卫艺办公助手")
    st.success("● AI 核心引擎已连接") 
    
    st.markdown("---")
    
    mode = st.radio("功能切换：", [
        "✨ 体卫艺简报助手",
        "📝 领导公务单自动生成器", 
        "🔍 龙华学校查号台"
    ])
    
    st.markdown("---")
    st.info("""
    **💡 助手功能说明：**
    
    1. **简报助手**：
       智能润色会议简报
    
    2. **公务单生成**：
       语音口语 → 规范公文Word
       
    3. **学校查号台**：
       全区通讯录一键查询
    """)
    st.caption("维护者：孙沛 | 龙华区教育局体卫艺专用")
    
    st.write("") # Spacer
    if st.button("🔒 退出并锁定系统"):
        st.session_state.contacts_authenticated = False
        st.session_state.parseddata_doc = None
        st.rerun()

# ----------------- 模块一：体卫艺简报助手 -----------------
if mode == "✨ 体卫艺简报助手":
    st.caption("↖️ **导航提示：** 点击左上角 **>** 图标打开菜单，可切换至其他功能")
    st.markdown("# ✨ 体卫艺简报助手")
    st.caption("@Technical Support Provided by Peipei")
    
    with st.container(border=True):
        # 助手介绍
        st.markdown("""
        <div style='background-color: #f0f7ff; padding: 1rem; border-radius: 8px; margin-bottom: 1rem; border-left: 4px solid #667eea;'>
            <p style='margin: 0; line-height: 1.6; color: #333;'>
                🤖 你好，我是擅长将杂乱信息转化为规范政务简讯的小助手，能为你打造高质量的体卫艺相关简报。
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.info("👇 **请直接发送：会议通知 + 参会名单 + 杂乱语音稿**")
        st.write("") 
        
        # Coze 商店链接
        DEEPSEEK_LINK = "https://www.coze.cn/store/agent/7587031903597985832?from=store_search_suggestion&bid=6ilacph8g3009"
        
        st.link_button("🚀 点击启动", DEEPSEEK_LINK)
        
        st.write("")
        st.markdown("""
        <small style='color:gray'>
        💡 <b>使用说明：</b><br>
        1. 点击按钮将跳转至体卫艺专用 AI 页面。<br>
        2. 支持超长文本处理与 DeepSeek 深度思考。<br>
        3. <b>无需配置 Key，永久免费使用。</b>
        </small>
        """, unsafe_allow_html=True)

# ----------------- 模块二：领导公务单自动生成器 -----------------
elif mode == "📝 领导公务单自动生成器":
    
    # 导航提示 (针对手机端用户不明显的问题)
    st.caption("↖️ **导航提示：** 点击左上角 **>** 图标打开菜单，可切换至「学校查号台」")
    
    # 使用容器包裹标题区域，打造卡片感
    st.markdown("# 📋 体卫艺领导公务单自动生成器")
    st.caption("Technical Support Provided by Peipei")
    
    # 蓝色提示框 - 提示语
    st.info("""
    **💡 智能提示：** 请一次性说清：时间、地点、会议名称、人数、对接人、领导、参加部门及议程。
    
    **🗣️ 参考范例：** “明天上午10点在二楼多功能厅有个生涯教育座谈会，大概20人，孙沛对接，1小时，邀请灵芝主任参加。”
    """)

    # --- 第一步：输入与润色 ---
    if st.session_state.step == 1:
        # 输入区卡片
        with st.container(border=True):
            st.subheader("1️⃣ 描述活动信息")
            st.caption("支持直接粘贴语音转文字的内容，AI 将自动提取要素。")
            
            user_input = st.text_area(
                "请在此输入...", 
                height=150, 
                placeholder="请点击此处粘贴或输入内容...", 
                key="input_doc", 
                label_visibility="collapsed"
            )
        
        st.write("") # 间距
        if st.button("✨ 立即智能填表并生成 Word", type="primary", use_container_width=True):
            if not user_input:
                st.warning("⚠️ 内容不能为空，请输入活动描述。")
            else:
                client = OpenAI(api_key=MY_API_KEY, base_url=BASE_URL)
                st.session_state.original_input = user_input
                
                # 获取当前日期用于计算相对时间
                current_date_str = datetime.now().strftime("%Y年%m月%d日")
                weekday = datetime.now().strftime("%w")
                
                with st.spinner("🤖 正在解析要素并润色公文语言..."):
                    
                    # 标准人名库（用于纠正语音转文字的谐音错误）
                    name_corrections = {
                        "林芝": "杨灵芝", "杨林芝": "杨灵芝",
                        "陈海湾": "陈海万", "陈海完": "陈海万",
                        "尹泽力": "尹泽利", "尹则利": "尹泽利",
                        "文量方": "文良方", "温良方": "文良方",
                        "刘兵": "刘冰",
                        "梁永育": "梁永誉",
                        "方梦仪": "方梦懿"
                    }
                    
                    full_prompt = f"""
                    你现在是龙华教育局资深笔杆子。请根据以下用户的大白话描述，解析出公文要素，并对【理由背景】和【议程】部分进行专业润色。
                    
                    【当前日期参考】：今天是 {current_date_str} (星期{weekday})。
                    【用户输入】：{user_input}
                    
                    【标准人名库】（请优先匹配）：
                    杨灵芝、尹泽利、文良方、孙沛、刘冰、杨帆、陈海万、路旭阳、王轩、王燕、李桂情、甘月琴、方梦懿、吴正光、李长生、梁永誉、刘喜菊
                    
                    【解析与润色要求】：
                    1. **人名纠错**：如果用户输入的人名与标准人名库相似（如"林芝"应为"杨灵芝"，"陈海湾"应为"陈海万"），请自动纠正为标准名字。
                    2. **content (理由背景)**：将用户的背景描述转化为"为落实...要求，推进...发展"等公文规范用语。只有动宾结构和语序调整，严禁杜撰。
                    3. **agenda (详细议程)**：**固定输出以下三项，顺序不可变**：["专题汇报", "座谈交流", "领导讲话"]。**严禁添加、删除或修改这三项**，无论用户输入什么。
                    4. **time (时间)**：必须将"明天"、"后天"、"周三"等相对时间**计算为具体的年月日**（格式：YYYY年MM月DD日 HH:MM）。禁止直接写"明天"或"下周"。
                    5. **duration (时长)**：统一计算为"X小时"或"X.5小时"（如1.5小时），**不要用分钟**。
                    6. **contact (公务对接人)**：提取人名（优先从标准人名库匹配），若无则默认为"孙沛"。
                    7. **dist_leader (区领导)** / **bur_leader (局领导)**：准确提取拟请出席的领导职务/姓名（如"灵芝主任"应识别为"杨灵芝"）。**严禁添加"教育发展中心"等部门前缀**，直接写姓名加职务即可。
                    8. **others (参加单位)**：提取建议参加的部门或单位。
                    9. **其他字段**：title(活动名称), place(地点), num(人数), projector(投影仪: ☑是/☐否)。
                    
                    必须以 JSON 格式严格输出，包含以下字段：
                    title, content, agenda, time, place, num, contact, projector, duration, dist_leader, bur_leader, others。
                    """
                    
                    try:
                        chat_completion = client.chat.completions.create(
                            model="Qwen/Qwen2.5-7B-Instruct", 
                            messages=[{"role": "user", "content": full_prompt}], 
                            response_format={'type': 'json_object'},
                            timeout=30 
                        )
                        result = json.loads(chat_completion.choices[0].message.content)
                        
                        # 字段健壮性处理
                        required_fields = ["title", "content", "agenda", "time", "place", "num", "contact"]
                        for field in required_fields:
                            if field not in result:
                                result[field] = ""
                        
                        st.session_state.parseddata_doc = result
                        st.session_state.step = 2  # 跳到确认表单
                        st.rerun()
                        
                    except json.JSONDecodeError:
                         st.error("❌ AI 解析返回格式有误，请尝试补充细节后重试。")
                    except TimeoutError:
                        st.error("⏱️ 请求超时，网络可能较慢。请稍后重试或简化输入内容。")
                    except Exception as e:
                        st.error(f"❌ 解析出错：{str(e)}")

    # --- 第二步：确认所有字段 ---
    elif st.session_state.step == 2 and st.session_state.parseddata_doc:
        d = st.session_state.parseddata_doc
        
        # 预览区卡片
        with st.container(border=True):
            st.subheader("2️⃣ 核心要素预览与微调")
            st.markdown("**📌 申报部门：体卫艺劳科**") 
            
            t = st.text_input("📝 政务活动名称", d.get("title", ""))
            c = st.text_area("📄 政务活动申请理由、背景", d.get("content", ""), height=100)
            
            # 处理 agenda
            agenda_val = d.get("agenda", "")
            if isinstance(agenda_val, list):
                agenda_val = "\n".join([f"{i+1}. {item}" for i, item in enumerate(agenda_val)])
            if not agenda_val:
                agenda_val = "1. 专题汇报\n2. 座谈交流\n3. 领导讲话"
            a = st.text_area("📋 议程", agenda_val, height=120)
            
            st.divider() # 分割线
            
            col1, col2 = st.columns(2)
            with col1:
                tm = st.text_input("⏰ 时间", d.get("time", ""))
                
                # 确保时长有单位
                duration_val = d.get("duration", "1小时")
                duration_val = str(duration_val) if duration_val else "1小时"
                if "小时" not in duration_val:
                    duration_val = f"{duration_val}小时"
                dr = st.text_input("⏳ 会议时长", duration_val)
                
            with col2:
                st.text_input("🚫 时间调整", "不可调整", disabled=True) 
                ct = st.text_input("👤 公务对接人", d.get("contact", "孙沛"))

            col3, col4, col5 = st.columns([2, 1, 1])
            with col3:
                pl = st.text_input("📍 地点", d.get("place", ""))
            with col4:
                nm = st.text_input("👥 人数", d.get("num", ""))
            with col5:
                pj = st.selectbox("📽️ 投影仪", ["☑使用", "☐不使用"], index=0 if "是" in str(d.get("projector")) else 1)
            
            st.divider()
            st.markdown("**👑 领导出席**")
            dist_l = st.text_input("1. 拟请出席的区领导", d.get("dist_leader", ""))
            bur_l = st.text_input("2. 拟请办公室协调出席的局领导", d.get("bur_leader", ""))
            
            st.divider()
            oth = st.text_input("🏛️ 建议参加单位(部门)", d.get("others") or "体卫艺劳科")
            
            st.caption("ℹ️ 说明：此表请于政务活动前一周星期四下班前交办公室登记汇总。")

        col_final_back, col_final_down = st.columns([1, 2])
        with col_final_back:
            if st.button("⬅️ 返回修改"):
                 st.session_state.step = 1
                 st.rerun()

        with col_final_down:
            try:
                final_data = {
                    "title": t, "content": c, "agenda": a, "time": tm, 
                    "duration": dr, "place": pl, "num": nm, "contact": ct, 
                    "projector": pj, "dist_leader": dist_l, "bur_leader": bur_l, "others": oth
                }
                tpl = DocxTemplate("template.docx")
                tpl.render(final_data)
                bio = io.BytesIO()
                tpl.save(bio)
                
                mmdd = datetime.now().strftime("%m%d")
                leader_name = bur_l.strip() if bur_l.strip() else (dist_l.strip() if dist_l.strip() else "领导")
                leader_name = leader_name.split('、')[0] if '、' in leader_name else leader_name
                filename = f"{mmdd}_{leader_name}_体卫艺劳科_{t}.docx"
                
                # 注入自定义样式，让下载按钮在不改变原生type的情况下变色
                st.markdown("""
                <style>
                    /* 定位最后一个按钮（通常是下载按钮，因为返回按钮在它前面） */
                    div.stButton > button:nth-last-child(1) {
                         background-color: #2e7d32 !important; /* 绿色 */
                         color: white !important;
                         border: none !important;
                    }
                </style>
                """, unsafe_allow_html=True)
                
                # 核心：直接使用原生按钮，触发微信的系统拦截机制 - 绝对不动
                st.download_button(
                    label="💾 确认无误，导出 Word",
                    data=bio.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"生成失败：{e}")

# ----------------- 模块三：龙华学校查号台 -----------------
elif mode == "🔍 龙华学校查号台":
    # 导航提示
    st.caption("↖️ **导航提示：** 点击左上角 **>** 图标打开菜单，可切换至其他功能")
    
    st.markdown("### 🔍 龙华学校查号台")
    st.caption("全区学校通讯录快速查询系统")
    
    if not st.session_state.contacts_authenticated:
        # 登录卡片
        with st.container(border=True):
            st.info("🔒 内部数据访问受限")
            pwd = st.text_input("请输入授权密码", type="password", help="请向管理员获取密码")
            if st.button("验证登录", type="primary", use_container_width=True):
                if pwd == CONTACT_PASSWORD:
                    st.session_state.contacts_authenticated = True
                    st.rerun()
                else:
                    st.error("❌ 密码错误，请重试。")
        st.stop()

    @st.cache_data
    def load_contacts():
        try:
            return pd.read_csv('contacts.csv', encoding='utf-8-sig').fillna('无')
        except:
            return pd.read_csv('contacts.csv', encoding='gbk').fillna('无')

    df = load_contacts()
    
    # 搜索框卡片
    with st.container(border=True):
        q = st.text_input("🔎 快速搜索", placeholder="输入学校名或人名关键词（如：龙华中学 或 张三）...")
        
    if q:
        mask = df.apply(lambda r: any(q.lower() in str(v).lower() for v in r.values), axis=1)
        st.write(f"📊 搜索结果：找到 {len(df[mask])} 条记录")
        st.dataframe(df[mask], use_container_width=True, hide_index=True)
    else:
        st.caption("👆 在上方输入关键词开始搜索，支持模糊匹配。")
