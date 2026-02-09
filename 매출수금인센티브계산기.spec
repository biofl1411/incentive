# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['intest.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 필요한 리소스 파일 추가 (예: 설정 파일, 폰트 등)
        ('incentive_calculator_settings.json', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'pandas',
        'numpy',
        'openpyxl',
        'dateutil',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 한글 폰트 지원 (시스템 폰트 경로 추가)
import os
font_path = os.path.join(os.environ['WINDIR'], 'Fonts')
a.datas += [
    (os.path.join(font_path, font), os.path.join('fonts', font), 'DATA') 
    for font in os.listdir(font_path) 
    if font.endswith(('.ttf', '.ttc'))
]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='매출수금인센티브계산기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI 앱이므로 콘솔 비활성화
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)