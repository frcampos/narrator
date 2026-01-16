# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('idiomas.py', '.'), ('tts_engine.py', '.'), ('pptx_handler.py', '.'), ('video_generator.py', '.'), ('tradutor.py', '.')]
binaries = []
hiddenimports = ['customtkinter', 'edge_tts', 'pptx', 'pydub', 'PIL', 'moviepy', 'moviepy.editor', 'moviepy.video.io.VideoFileClip', 'moviepy.audio.io.AudioFileClip', 'imageio', 'imageio_ffmpeg', 'proglog', 'deep_translator', 'pygame']
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('moviepy')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PPTX_Narrator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
app = BUNDLE(
    exe,
    name='PPTX_Narrator.app',
    icon=None,
    bundle_identifier=None,
)
