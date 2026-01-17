# -*- mode: python ; coding: utf-8 -*-
# Mac 版本打包配置文件

block_cipher = None

a = Analysis(
    ['pdf_page_remover.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('HYQiHeiClassic-55S.ttf', '.'),
        ('HYQiHeiClassic-60S.ttf', '.'),
        ('HYQiHeiClassic-70S.ttf', '.'),
        ('newlogo.png', '.'),
        ('newlogo2.jpeg', '.'),
    ],
    hiddenimports=[
        # PyQt5相关
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.sip',
        # PyPDF2相关
        'PyPDF2',
        'PyPDF2._reader',
        'PyPDF2._writer',
        # PyMuPDF相关
        'fitz',
        # matplotlib相关（用于信用分可视化）
        'matplotlib',
        'matplotlib.backends',
        'matplotlib.backends.backend_agg',
        'matplotlib.figure',
        'matplotlib.font_manager',
        'matplotlib.pyplot',
        'matplotlib.patches',
        'matplotlib.text',
        # numpy相关（matplotlib依赖）
        'numpy',
        # requests相关
        'requests',
        'urllib3',
        'certifi',
        # 自定义模块
        'credit_score_visualizer',
        # PIL相关（某些情况下需要）
        'PIL',
        'PIL._tkinter_finder',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # ========== AI/ML / 深度学习 / 视觉相关大型库（本项目不需要）==========
        'sklearn', 'scikit-learn',
        'scipy',
        'pandas',
        'seaborn',
        'torch', 'torchvision', 'torchaudio',
        'pytorch-lightning', 'torchmetrics', 'torchsde', 'torchdiffeq',
        'tensorflow', 'tensorflow-intel', 'tensorflow-estimator',
        'tensorflow-io-gcs-filesystem', 'keras', 'tf-keras-nightly',
        'jax', 'jaxlib', 'ml_dtypes',
        'transformers', 'diffusers', 'accelerate', 'xformers',
        'taming-transformers', 'taming-transformers-rom1504',
        'timm', 'lpips',
        'huggingface-hub', 'safetensors', 'sentencepiece', 'tokenizers',
        'open-clip-torch',
        'clean-fid', 'pytorch-fid',
        'basicsr', 'gfpgan', 'realesrgan',
        'facexlib',
        'opencv-python', 'opencv-contrib-python', 'cv2',
        'mediapipe',
        'ultralytics', 'ultralytics-thop',
        'tqdm',

        # ========== LLM / LangChain / OpenAI 等（本项目不需要）==========
        'openai', 'tiktoken',
        'langchain', 'langchain-core', 'langchain-community',
        'langchain-classic', 'langchain-openai', 'langchain-text-splitters',
        'langgraph', 'langgraph-checkpoint', 'langgraph-prebuilt', 'langgraph-sdk',
        'langsmith',

        # ========== Jupyter / IPython / Notebook 相关 ==========
        'IPython', 'ipython',
        'jupyter', 'jupyterlab', 'jupyterlab_server',
        'jupyter_client', 'jupyter-console', 'jupyter-core',
        'jupyter_server', 'jupyter-events', 'jupyter_server_terminals',
        'jupyterlab-widgets', 'widgetsnbextension',
        'notebook', 'notebooks', 'notebook_shim',
        'ipykernel',

        # ========== Web / API / 前端服务框架 ==========
        'fastapi', 'starlette', 'uvicorn',
        'Flask', 'flask', 'Flask-Cors',
        'dashscope',

        # ========== 其他GUI框架 ==========
        'tkinter', 'Tkinter',
        'wx', 'wxpython',

        # ========== 测试框架 ==========
        'pytest', 'nose',
        'doctest',
        # 注意：不移除 'unittest'，因为某些依赖（如 PyQt5）可能在运行时需要它

        # ========== 标准库 / 工具中运行时不需要的部分 ==========
        'xmlrpc', 'xmlrpc.client', 'xmlrpc.server',
        'pydoc',
        'distutils',
        'setuptools',
        'pip',
        'wheel',

        # ========== 开发 / 调试 / 性能分析 ==========
        'pdb',
        'profile', 'pstats',
        'cProfile',

        # ========== 其他不需要的库 / 测试模块 ==========
        'matplotlib.tests',
        'numpy.tests',
        'PyQt5.tests',
    ],
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PDF处理器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Mac GUI应用不显示控制台
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,  # 使用运行环境的默认架构
    codesign_identity=None,  # 如果需要代码签名，填写开发者ID，例如: 'Developer ID Application: Your Name'
    entitlements_file=None,  # 如果需要特殊权限，指定entitlements文件路径
    icon=None,  # Mac图标文件路径，格式为 .icns，例如: 'icon.icns'
)

# Mac 应用需要 COLLECT 来创建 .app 目录结构
app = BUNDLE(
    exe,
    name='PDF处理器.app',
    icon=None,  # Mac图标文件路径，格式为 .icns
    bundle_identifier='com.yourcompany.pdfprocessor',  # 应用标识符，建议修改为你的唯一标识
    version=None,  # 应用版本号，例如: '1.0.0'
    info_plist={
        'NSHighResolutionCapable': 'True',  # 支持高分辨率显示
        'NSRequiresAquaSystemAppearance': 'False',  # 支持深色模式
        'LSMinimumSystemVersion': '10.15',  # 最低支持的 macOS 版本 (Catalina)
    },
)
