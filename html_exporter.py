#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
html_exporter.py - Exportador de apresentações para HTML5
Versão 1.0 - Suporta karaoke, múltiplos layouts e responsivo

Modos de layout:
- separated: Área separada para texto (slide em cima, texto em baixo)
- overlay: Texto sobreposto no slide
- main: Texto como corpo principal (slide pequeno)
- textonly: Apenas texto (sem slide)
"""

import os
import json
import base64
import shutil
from pathlib import Path
from typing import Optional, List, Dict, Any, Callable
from dataclasses import dataclass, field, asdict
from datetime import datetime


@dataclass
class ConfigLayoutHTML:
    """Configuração de layout"""
    mode: str = "separated"  # separated, overlay, main, textonly
    show_slide: bool = True
    slide_size: str = "medium"  # large, medium, small, thumbnail


@dataclass
class ConfigAudioHTML:
    """Configuração de áudio"""
    enabled: bool = True
    language: str = "original"  # original, translated, both
    autoplay: bool = True


@dataclass
class ConfigKaraokeHTML:
    """Configuração de karaoke"""
    enabled: bool = True
    text_language: str = "same"  # same, original, translated
    highlight_color: str = "#FFFF00"
    highlight_opacity: float = 0.4  # Opacidade do marcador (0.1 a 1.0)
    scroll_mode: str = "scroll"  # scroll, page


@dataclass
class ConfigNavigationHTML:
    """Configuração de navegação"""
    show_controls: bool = True
    show_progress: bool = True
    keyboard_nav: bool = True
    slide_index: bool = False
    touch_gestures: bool = True


@dataclass
class ConfigAppearanceHTML:
    """Configuração de aparência"""
    theme: str = "dark"  # dark, light, highcontrast
    font_size: str = "medium"  # small, medium, large, xlarge
    font_family: str = "sans-serif"


@dataclass
class ConfigExportHTML:
    """Configuração completa de exportação HTML"""
    format: str = "folder"  # folder, single
    resources: str = "embedded"  # embedded, cdn
    layout: ConfigLayoutHTML = field(default_factory=ConfigLayoutHTML)
    audio: ConfigAudioHTML = field(default_factory=ConfigAudioHTML)
    karaoke: ConfigKaraokeHTML = field(default_factory=ConfigKaraokeHTML)
    navigation: ConfigNavigationHTML = field(default_factory=ConfigNavigationHTML)
    appearance: ConfigAppearanceHTML = field(default_factory=ConfigAppearanceHTML)
    
    def to_dict(self) -> dict:
        """Converte para dicionário"""
        return {
            "format": self.format,
            "resources": self.resources,
            "layout": asdict(self.layout),
            "audio": asdict(self.audio),
            "karaoke": asdict(self.karaoke),
            "navigation": asdict(self.navigation),
            "appearance": asdict(self.appearance)
        }


@dataclass
class SlideDataHTML:
    """Dados de um slide para HTML"""
    id: int
    image_path: str = ""
    audio_original: str = ""
    audio_translated: str = ""
    duration: float = 0.0
    text_original: str = ""
    text_translated: str = ""
    words: List[Dict] = field(default_factory=list)


class HTMLExporter:
    """Exportador de apresentações para HTML5"""
    
    def __init__(self):
        self.config = ConfigExportHTML()
        self.slides: List[SlideDataHTML] = []
        self.title = "Apresentação"
        self._progress_callback: Optional[Callable] = None
    
    def set_progress_callback(self, callback: Callable):
        """Define callback para reportar progresso"""
        self._progress_callback = callback
    
    def _report_progress(self, current: int, total: int, message: str = ""):
        """Reporta progresso"""
        if self._progress_callback:
            self._progress_callback(current, total, message)
    
    def export(self, output_path: str, slides_data: List[SlideDataHTML], 
               title: str = "Apresentação") -> bool:
        """
        Exporta apresentação para HTML5.
        
        Args:
            output_path: Caminho de saída (pasta ou ficheiro .html)
            slides_data: Lista de dados dos slides
            title: Título da apresentação
        
        Returns:
            True se sucesso
        """
        self.slides = slides_data
        self.title = title
        
        try:
            if self.config.format == "folder":
                return self._export_folder(output_path)
            else:
                return self._export_single(output_path)
        except Exception as e:
            print(f"Erro ao exportar HTML: {e}")
            return False
    
    def _export_folder(self, output_path: str) -> bool:
        """Exporta como pasta com ficheiros separados"""
        self._report_progress(0, 100, "A criar estrutura de pastas...")
        
        # Criar estrutura de pastas
        base_path = Path(output_path)
        assets_path = base_path / "assets"
        slides_path = assets_path / "slides"
        audio_path = assets_path / "audio"
        css_path = assets_path / "css"
        js_path = assets_path / "js"
        data_path = base_path / "data"
        
        for path in [base_path, assets_path, slides_path, audio_path, 
                     css_path, js_path, data_path]:
            path.mkdir(parents=True, exist_ok=True)
        
        self._report_progress(10, 100, "A copiar recursos...")
        
        # Copiar slides e áudios
        slides_json = []
        total_slides = len(self.slides)
        
        for i, slide in enumerate(self.slides):
            self._report_progress(
                10 + int(50 * i / total_slides), 100,
                f"A processar slide {slide.id}..."
            )
            
            slide_data = {
                "id": slide.id,
                "image": "",
                "audio": {"original": "", "translated": ""},
                "duration": slide.duration,
                "text": {
                    "original": slide.text_original,
                    "translated": slide.text_translated
                },
                "words": slide.words
            }
            
            # Copiar imagem do slide
            if slide.image_path and os.path.exists(slide.image_path):
                ext = Path(slide.image_path).suffix
                dest_img = slides_path / f"slide_{slide.id:02d}{ext}"
                shutil.copy2(slide.image_path, dest_img)
                slide_data["image"] = f"assets/slides/slide_{slide.id:02d}{ext}"
            
            # Copiar áudio original
            if slide.audio_original and os.path.exists(slide.audio_original):
                ext = Path(slide.audio_original).suffix
                dest_audio = audio_path / f"slide_{slide.id:02d}{ext}"
                shutil.copy2(slide.audio_original, dest_audio)
                slide_data["audio"]["original"] = f"assets/audio/slide_{slide.id:02d}{ext}"
            
            # Copiar áudio traduzido
            if slide.audio_translated and os.path.exists(slide.audio_translated):
                ext = Path(slide.audio_translated).suffix
                dest_audio = audio_path / f"slide_{slide.id:02d}_trad{ext}"
                shutil.copy2(slide.audio_translated, dest_audio)
                slide_data["audio"]["translated"] = f"assets/audio/slide_{slide.id:02d}_trad{ext}"
            
            slides_json.append(slide_data)
        
        self._report_progress(60, 100, "A gerar CSS...")
        
        # Gerar CSS
        css_content = self._generate_css()
        with open(css_path / "player.css", 'w', encoding='utf-8') as f:
            f.write(css_content)
        
        self._report_progress(70, 100, "A gerar JavaScript...")
        
        # Gerar JavaScript
        js_content = self._generate_js()
        with open(js_path / "player.js", 'w', encoding='utf-8') as f:
            f.write(js_content)
        
        self._report_progress(80, 100, "A preparar dados...")
        
        # Preparar dados da apresentação
        presentation_data = {
            "version": "1.0",
            "title": self.title,
            "created": datetime.now().isoformat(),
            "config": self.config.to_dict(),
            "slides": slides_json
        }
        
        # Também guardar JSON separado (para referência/debug)
        with open(data_path / "presentation.json", 'w', encoding='utf-8') as f:
            json.dump(presentation_data, f, ensure_ascii=False, indent=2)
        
        self._report_progress(90, 100, "A gerar HTML...")
        
        # Gerar HTML principal COM DADOS EMBEDADOS (evita CORS)
        html_content = self._generate_html_folder(
            css_path="assets/css/player.css",
            js_path="assets/js/player.js",
            presentation_data=presentation_data
        )
        
        with open(base_path / "index.html", 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        self._report_progress(100, 100, "Exportação concluída!")
        return True
    
    def _export_single(self, output_path: str) -> bool:
        """Exporta como ficheiro HTML único"""
        self._report_progress(0, 100, "A preparar ficheiro único...")
        
        # Preparar dados dos slides com base64
        slides_json = []
        total_slides = len(self.slides)
        
        for i, slide in enumerate(self.slides):
            self._report_progress(
                int(40 * i / total_slides), 100,
                f"A converter slide {slide.id}..."
            )
            
            slide_data = {
                "id": slide.id,
                "image": "",
                "audio": {"original": "", "translated": ""},
                "duration": slide.duration,
                "text": {
                    "original": slide.text_original,
                    "translated": slide.text_translated
                },
                "words": slide.words
            }
            
            # Converter imagem para base64
            if slide.image_path and os.path.exists(slide.image_path):
                slide_data["image"] = self._file_to_base64(slide.image_path, "image")
            
            # Converter áudios para base64
            if slide.audio_original and os.path.exists(slide.audio_original):
                slide_data["audio"]["original"] = self._file_to_base64(
                    slide.audio_original, "audio"
                )
            
            if slide.audio_translated and os.path.exists(slide.audio_translated):
                slide_data["audio"]["translated"] = self._file_to_base64(
                    slide.audio_translated, "audio"
                )
            
            slides_json.append(slide_data)
        
        self._report_progress(50, 100, "A gerar HTML...")
        
        # Preparar dados inline
        presentation_data = {
            "version": "1.0",
            "title": self.title,
            "created": datetime.now().isoformat(),
            "config": self.config.to_dict(),
            "slides": slides_json
        }
        
        # Gerar CSS e JS
        css_content = self._generate_css()
        js_content = self._generate_js()
        
        self._report_progress(70, 100, "A compilar ficheiro...")
        
        # Gerar HTML com tudo inline
        html_content = self._generate_html_single(
            css_content, js_content, presentation_data
        )
        
        # Garantir extensão .html
        if not output_path.endswith('.html'):
            output_path += '.html'
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        self._report_progress(100, 100, "Exportação concluída!")
        return True
    
    def _file_to_base64(self, filepath: str, media_type: str) -> str:
        """Converte ficheiro para data URI base64"""
        ext = Path(filepath).suffix.lower()
        
        mime_types = {
            # Imagens
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.png': 'image/png',
            '.gif': 'image/gif',
            '.webp': 'image/webp',
            # Áudio
            '.mp3': 'audio/mpeg',
            '.wav': 'audio/wav',
            '.ogg': 'audio/ogg',
            '.m4a': 'audio/mp4',
        }
        
        mime = mime_types.get(ext, f'{media_type}/*')
        
        with open(filepath, 'rb') as f:
            data = base64.b64encode(f.read()).decode('utf-8')
        
        return f"data:{mime};base64,{data}"
    
    def _generate_css(self) -> str:
        """Gera o CSS do player"""
        theme = self.config.appearance.theme
        layout = self.config.layout.mode
        
        # Variáveis de tema
        themes = {
            "dark": {
                "bg_primary": "#1a1a2e",
                "bg_secondary": "#16213e",
                "bg_tertiary": "#0f0f1a",
                "text_primary": "#ffffff",
                "text_secondary": "#a0a0a0",
                "accent": "#e94560",
                "highlight": self.config.karaoke.highlight_color,
                "overlay_bg": "rgba(0, 0, 0, 0.85)"
            },
            "light": {
                "bg_primary": "#ffffff",
                "bg_secondary": "#f5f5f5",
                "bg_tertiary": "#e0e0e0",
                "text_primary": "#333333",
                "text_secondary": "#666666",
                "accent": "#0066cc",
                "highlight": self.config.karaoke.highlight_color,
                "overlay_bg": "rgba(255, 255, 255, 0.9)"
            },
            "highcontrast": {
                "bg_primary": "#000000",
                "bg_secondary": "#000000",
                "bg_tertiary": "#1a1a1a",
                "text_primary": "#ffffff",
                "text_secondary": "#ffff00",
                "accent": "#00ff00",
                "highlight": "#00ffff",
                "overlay_bg": "rgba(0, 0, 0, 0.95)"
            }
        }
        
        t = themes.get(theme, themes["dark"])
        
        # Tamanhos de fonte
        font_sizes = {
            "small": {"base": "14px", "karaoke": "1.2rem", "karaoke_main": "1.8rem"},
            "medium": {"base": "16px", "karaoke": "1.5rem", "karaoke_main": "2.2rem"},
            "large": {"base": "18px", "karaoke": "1.8rem", "karaoke_main": "2.8rem"},
            "xlarge": {"base": "20px", "karaoke": "2.2rem", "karaoke_main": "3.5rem"}
        }
        
        fs = font_sizes.get(self.config.appearance.font_size, font_sizes["medium"])
        
        # Converter cor hex para rgba com opacidade
        highlight_color = t["highlight"]
        highlight_opacity = self.config.karaoke.highlight_opacity
        
        # Converter hex (#RRGGBB) para rgba
        if highlight_color.startswith('#') and len(highlight_color) == 7:
            r = int(highlight_color[1:3], 16)
            g = int(highlight_color[3:5], 16)
            b = int(highlight_color[5:7], 16)
            highlight_rgba = f"rgba({r}, {g}, {b}, {highlight_opacity})"
        else:
            highlight_rgba = highlight_color
        
        css = f'''/* PPTX Narrator - HTML5 Player CSS */
/* Gerado automaticamente - v1.0 */

:root {{
    --bg-primary: {t["bg_primary"]};
    --bg-secondary: {t["bg_secondary"]};
    --bg-tertiary: {t["bg_tertiary"]};
    --text-primary: {t["text_primary"]};
    --text-secondary: {t["text_secondary"]};
    --accent: {t["accent"]};
    --highlight: {t["highlight"]};
    --highlight-bg: {highlight_rgba};
    --overlay-bg: {t["overlay_bg"]};
    --font-base: {fs["base"]};
    --font-karaoke: {fs["karaoke"]};
    --font-karaoke-main: {fs["karaoke_main"]};
    --font-family: {self.config.appearance.font_family}, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}}

* {{
    margin: 0;
    padding: 0;
    box-sizing: border-box;
}}

html, body {{
    width: 100%;
    height: 100%;
    overflow: hidden;
    font-family: var(--font-family);
    font-size: var(--font-base);
    background: var(--bg-primary);
    color: var(--text-primary);
}}

/* ===== CONTAINER PRINCIPAL ===== */
.player-container {{
    width: 100%;
    height: 100%;
    display: flex;
    flex-direction: column;
}}

/* ===== ÁREA DO SLIDE ===== */
.slide-area {{
    flex: 1;
    display: flex;
    flex-direction: column;
    position: relative;
    overflow: hidden;
}}

.slide-image-container {{
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    background: var(--bg-tertiary);
    position: relative;
    overflow: hidden;
}}

.slide-image {{
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
}}

/* ===== LAYOUTS ===== */

/* Layout: Área Separada */
.layout-separated .slide-image-container {{
    flex: 0 0 70%;
}}

.layout-separated .karaoke-area {{
    flex: 0 0 30%;
    background: var(--bg-secondary);
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 1rem;
}}

/* Layout: Sobreposto */
.layout-overlay .slide-image-container {{
    flex: 1;
}}

.layout-overlay .karaoke-area {{
    position: absolute;
    bottom: 0;
    left: 0;
    right: 0;
    background: var(--overlay-bg);
    padding: 1.5rem;
    backdrop-filter: blur(10px);
}}

/* Layout: Corpo Principal */
.layout-main {{
    flex-direction: column;
}}

.layout-main .slide-image-container {{
    flex: 0 0 30%;
    max-height: 30%;
}}

.layout-main .slide-image {{
    max-height: 100%;
}}

.layout-main .karaoke-area {{
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 2rem;
    background: var(--bg-secondary);
}}

.layout-main .karaoke-text {{
    font-size: var(--font-karaoke-main);
    line-height: 1.6;
}}

/* Layout: Só Texto */
.layout-textonly .slide-image-container {{
    display: none;
}}

.layout-textonly .karaoke-area {{
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 3rem;
}}

.layout-textonly .karaoke-text {{
    font-size: var(--font-karaoke-main);
    line-height: 1.8;
    max-width: 900px;
}}

/* ===== KARAOKE ===== */
.karaoke-area {{
    text-align: center;
}}

.karaoke-text {{
    font-size: var(--font-karaoke);
    line-height: 1.5;
    color: var(--text-primary);
}}

.karaoke-word {{
    display: inline;
    transition: all 0.1s ease;
    padding: 0.1em 0.15em;
    border-radius: 4px;
}}

.karaoke-word.active {{
    background: var(--highlight-bg);
    color: var(--text-primary);
    font-weight: 600;
}}

.karaoke-word.past {{
    color: var(--text-secondary);
}}

/* ===== CONTROLOS ===== */
.controls-area {{
    background: var(--bg-secondary);
    padding: 0.8rem 1rem;
    display: flex;
    flex-direction: column;
    gap: 0.5rem;
    border-top: 1px solid var(--bg-tertiary);
}}

.progress-bar {{
    width: 100%;
    height: 6px;
    background: var(--bg-tertiary);
    border-radius: 3px;
    cursor: pointer;
    position: relative;
}}

.progress-bar-fill {{
    height: 100%;
    background: var(--accent);
    border-radius: 3px;
    width: 0%;
    transition: width 0.1s linear;
}}

.controls-buttons {{
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 1rem;
}}

.control-btn {{
    background: transparent;
    border: none;
    color: var(--text-primary);
    font-size: 1.5rem;
    cursor: pointer;
    padding: 0.5rem;
    border-radius: 50%;
    transition: all 0.2s;
    display: flex;
    align-items: center;
    justify-content: center;
    width: 48px;
    height: 48px;
}}

.control-btn:hover {{
    background: var(--bg-tertiary);
    color: var(--accent);
}}

.control-btn:active {{
    transform: scale(0.95);
}}

.control-btn.play-btn {{
    width: 56px;
    height: 56px;
    font-size: 1.8rem;
    background: var(--accent);
    color: white;
}}

.control-btn.play-btn:hover {{
    background: var(--accent);
    opacity: 0.9;
}}

.slide-info {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    font-size: 0.85rem;
    color: var(--text-secondary);
}}

.volume-control {{
    display: flex;
    align-items: center;
    gap: 0.5rem;
}}

.volume-slider {{
    width: 80px;
    height: 4px;
    -webkit-appearance: none;
    background: var(--bg-tertiary);
    border-radius: 2px;
    cursor: pointer;
}}

.volume-slider::-webkit-slider-thumb {{
    -webkit-appearance: none;
    width: 14px;
    height: 14px;
    background: var(--accent);
    border-radius: 50%;
    cursor: pointer;
}}

.language-selector {{
    display: flex;
    align-items: center;
    gap: 0.3rem;
}}

.lang-btn {{
    padding: 0.3rem 0.6rem;
    border: 1px solid var(--bg-tertiary);
    background: transparent;
    color: var(--text-secondary);
    border-radius: 4px;
    cursor: pointer;
    font-size: 0.75rem;
    transition: all 0.2s;
}}

.lang-btn.active {{
    background: var(--accent);
    color: white;
    border-color: var(--accent);
}}

/* ===== ÍNDICE DE SLIDES ===== */
.slide-index {{
    position: fixed;
    left: 0;
    top: 0;
    bottom: 0;
    width: 200px;
    background: var(--bg-secondary);
    transform: translateX(-100%);
    transition: transform 0.3s ease;
    z-index: 100;
    overflow-y: auto;
    padding: 1rem;
}}

.slide-index.open {{
    transform: translateX(0);
}}

.slide-index-item {{
    padding: 0.5rem;
    margin-bottom: 0.5rem;
    border-radius: 8px;
    cursor: pointer;
    transition: background 0.2s;
}}

.slide-index-item:hover {{
    background: var(--bg-tertiary);
}}

.slide-index-item.active {{
    background: var(--accent);
}}

.slide-index-thumb {{
    width: 100%;
    border-radius: 4px;
    margin-bottom: 0.3rem;
}}

/* ===== RESPONSIVO - TABLET ===== */
@media (max-width: 1024px) {{
    .layout-separated .slide-image-container {{
        flex: 0 0 65%;
    }}
    
    .layout-separated .karaoke-area {{
        flex: 0 0 35%;
    }}
    
    .layout-main .slide-image-container {{
        flex: 0 0 35%;
    }}
}}

/* ===== RESPONSIVO - MOBILE ===== */
@media (max-width: 768px) {{
    .layout-separated .slide-image-container {{
        flex: 0 0 55%;
    }}
    
    .layout-separated .karaoke-area {{
        flex: 0 0 45%;
        padding: 0.8rem;
    }}
    
    .layout-main .slide-image-container {{
        flex: 0 0 25%;
    }}
    
    .layout-main .karaoke-area {{
        padding: 1rem;
    }}
    
    .karaoke-text {{
        font-size: 1.2rem;
    }}
    
    .layout-main .karaoke-text,
    .layout-textonly .karaoke-text {{
        font-size: 1.6rem;
    }}
    
    .controls-buttons {{
        gap: 0.5rem;
    }}
    
    .control-btn {{
        width: 40px;
        height: 40px;
        font-size: 1.2rem;
    }}
    
    .control-btn.play-btn {{
        width: 48px;
        height: 48px;
        font-size: 1.4rem;
    }}
    
    .volume-control {{
        display: none;
    }}
    
    .slide-info {{
        font-size: 0.75rem;
    }}
}}

/* ===== RESPONSIVO - MOBILE PEQUENO ===== */
@media (max-width: 480px) {{
    .layout-separated .slide-image-container {{
        flex: 0 0 50%;
    }}
    
    .layout-separated .karaoke-area {{
        flex: 0 0 50%;
    }}
    
    .karaoke-text {{
        font-size: 1rem;
    }}
    
    .layout-main .karaoke-text,
    .layout-textonly .karaoke-text {{
        font-size: 1.3rem;
    }}
    
    .controls-area {{
        padding: 0.5rem;
    }}
}}

/* ===== FULLSCREEN ===== */
.player-container:fullscreen {{
    background: var(--bg-primary);
}}

.player-container:fullscreen .controls-area {{
    position: absolute;
    bottom: 0;
    left: 0;
    right: 0;
    opacity: 0;
    transition: opacity 0.3s;
}}

.player-container:fullscreen:hover .controls-area {{
    opacity: 1;
}}

/* ===== LOADING ===== */
.loading {{
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    text-align: center;
}}

.loading-spinner {{
    width: 40px;
    height: 40px;
    border: 3px solid var(--bg-tertiary);
    border-top-color: var(--accent);
    border-radius: 50%;
    animation: spin 1s linear infinite;
}}

@keyframes spin {{
    to {{ transform: rotate(360deg); }}
}}

/* ===== UTILITIES ===== */
.hidden {{
    display: none !important;
}}
'''
        return css
    
    def _generate_js(self) -> str:
        """Gera o JavaScript do player"""
        js = '''/* PPTX Narrator - HTML5 Player JavaScript */
/* Gerado automaticamente - v1.0 */

class PresentationPlayer {
    constructor() {
        this.currentSlide = 0;
        this.isPlaying = false;
        this.audioLanguage = 'original';
        this.presentationData = null;
        this.audioElement = null;
        this.karaokeInterval = null;
        this.wordTimings = [];
        this.currentWordIndex = 0;
        this.currentWordCount = 0;  // Para recalcular timings quando áudio carrega
        
        this.init();
    }
    
    async init() {
        // Carregar dados da apresentação
        await this.loadPresentationData();
        
        // Configurar elementos DOM (ANTES de applyConfig)
        this.setupDOM();
        
        // Aplicar configurações (DEPOIS de setupDOM)
        if (this.presentationData?.config) {
            this.applyConfig(this.presentationData.config);
        }
        
        // Configurar eventos
        this.setupEvents();
        
        // Carregar primeiro slide
        if (this.presentationData?.slides?.length > 0) {
            this.loadSlide(0);
        } else {
            console.error('Nenhum slide encontrado na apresentação');
        }
    }
    
    async loadPresentationData() {
        // Método 1: Dados já definidos globalmente (modo single file)
        if (window.PRESENTATION_DATA) {
            this.presentationData = window.PRESENTATION_DATA;
            console.log('Dados carregados de PRESENTATION_DATA');
            return;
        } 
        
        // Método 2: Dados embedados em script tag (modo pasta - evita CORS)
        const dataScript = document.getElementById('presentation-data');
        if (dataScript) {
            try {
                this.presentationData = JSON.parse(dataScript.textContent);
                console.log('Dados carregados de script embedado');
                return;
            } catch (e) {
                console.error('Erro ao parsear dados embedados:', e);
            }
        }
        
        // Método 3: Fetch externo (fallback - pode falhar com file://)
        try {
            const response = await fetch('data/presentation.json');
            this.presentationData = await response.json();
            console.log('Dados carregados via fetch');
        } catch (e) {
            console.error('Erro ao carregar dados via fetch:', e);
            // Criar dados vazios para evitar erros
            this.presentationData = { slides: [], config: {} };
        }
    }
    
    applyConfig(config) {
        const container = document.querySelector('.player-container');
        if (!container) return;
        
        // Aplicar layout
        container.className = 'player-container';
        if (config.layout?.mode) {
            container.classList.add(`layout-${config.layout.mode}`);
        }
        
        // Esconder slide se configurado
        if (config.layout && !config.layout.show_slide) {
            document.querySelector('.slide-image-container')?.classList.add('hidden');
        }
        
        // Configurar idioma de áudio
        if (config.audio?.language) {
            this.audioLanguage = config.audio.language === 'translated' ? 'translated' : 'original';
        }
        
        // Actualizar botões de idioma (agora this.elements existe)
        if (this.elements) {
            this.updateLanguageButtons();
        }
    }
    
    setupDOM() {
        this.elements = {
            slideImage: document.getElementById('slide-image'),
            karaokeText: document.getElementById('karaoke-text'),
            playBtn: document.getElementById('play-btn'),
            prevBtn: document.getElementById('prev-btn'),
            nextBtn: document.getElementById('next-btn'),
            progressBar: document.getElementById('progress-bar'),
            progressFill: document.getElementById('progress-fill'),
            currentTime: document.getElementById('current-time'),
            totalTime: document.getElementById('total-time'),
            slideCounter: document.getElementById('slide-counter'),
            volumeSlider: document.getElementById('volume-slider'),
            langOriginal: document.getElementById('lang-original'),
            langTranslated: document.getElementById('lang-translated'),
            fullscreenBtn: document.getElementById('fullscreen-btn')
        };
        
        // Criar elemento de áudio
        this.audioElement = new Audio();
        this.audioElement.addEventListener('timeupdate', () => this.onTimeUpdate());
        this.audioElement.addEventListener('ended', () => this.onAudioEnded());
        this.audioElement.addEventListener('loadedmetadata', () => this.onAudioLoaded());
    }
    
    setupEvents() {
        // Controlos
        this.elements.playBtn?.addEventListener('click', () => this.togglePlay());
        this.elements.prevBtn?.addEventListener('click', () => this.prevSlide());
        this.elements.nextBtn?.addEventListener('click', () => this.nextSlide());
        
        // Progresso
        this.elements.progressBar?.addEventListener('click', (e) => this.seekTo(e));
        
        // Volume
        this.elements.volumeSlider?.addEventListener('input', (e) => {
            this.audioElement.volume = e.target.value / 100;
        });
        
        // Idioma
        this.elements.langOriginal?.addEventListener('click', () => this.setLanguage('original'));
        this.elements.langTranslated?.addEventListener('click', () => this.setLanguage('translated'));
        
        // Fullscreen
        this.elements.fullscreenBtn?.addEventListener('click', () => this.toggleFullscreen());
        
        // Teclado
        document.addEventListener('keydown', (e) => this.onKeyDown(e));
        
        // Touch/Swipe
        this.setupTouchEvents();
    }
    
    setupTouchEvents() {
        let touchStartX = 0;
        let touchEndX = 0;
        
        const container = document.querySelector('.slide-area');
        
        container?.addEventListener('touchstart', (e) => {
            touchStartX = e.changedTouches[0].screenX;
        }, { passive: true });
        
        container?.addEventListener('touchend', (e) => {
            touchEndX = e.changedTouches[0].screenX;
            this.handleSwipe(touchStartX, touchEndX);
        }, { passive: true });
        
        // Tap para play/pause
        container?.addEventListener('click', (e) => {
            if (e.target.closest('.controls-area')) return;
            this.togglePlay();
        });
    }
    
    handleSwipe(startX, endX) {
        const threshold = 50;
        const diff = startX - endX;
        
        if (Math.abs(diff) > threshold) {
            if (diff > 0) {
                this.nextSlide();
            } else {
                this.prevSlide();
            }
        }
    }
    
    onKeyDown(e) {
        switch(e.code) {
            case 'Space':
                e.preventDefault();
                this.togglePlay();
                break;
            case 'ArrowLeft':
                this.prevSlide();
                break;
            case 'ArrowRight':
                this.nextSlide();
                break;
            case 'ArrowUp':
                this.audioElement.volume = Math.min(1, this.audioElement.volume + 0.1);
                break;
            case 'ArrowDown':
                this.audioElement.volume = Math.max(0, this.audioElement.volume - 0.1);
                break;
            case 'KeyL':
                this.toggleLanguage();
                break;
            case 'KeyF':
                this.toggleFullscreen();
                break;
            case 'KeyM':
                this.audioElement.muted = !this.audioElement.muted;
                break;
            default:
                // Números 1-9 para ir a slides específicos
                if (e.code.startsWith('Digit')) {
                    const num = parseInt(e.code.replace('Digit', ''));
                    if (num > 0 && num <= this.presentationData.slides.length) {
                        this.loadSlide(num - 1);
                    }
                }
        }
    }
    
    loadSlide(index) {
        if (!this.presentationData?.slides) return;
        
        const slides = this.presentationData.slides;
        if (index < 0 || index >= slides.length) return;
        
        this.currentSlide = index;
        const slide = slides[index];
        
        // Parar áudio anterior
        this.pause();
        
        // Carregar imagem
        if (this.elements.slideImage && slide.image) {
            this.elements.slideImage.src = slide.image;
        }
        
        // Carregar texto karaoke
        this.loadKaraokeText(slide);
        
        // Carregar áudio
        this.loadAudio(slide);
        
        // Actualizar contador
        this.updateSlideCounter();
        
        // Auto-play se configurado
        if (this.presentationData.config?.audio?.autoplay) {
            setTimeout(() => this.play(), 300);
        }
    }
    
    loadKaraokeText(slide) {
        if (!this.elements.karaokeText) return;
        
        const config = this.presentationData.config;
        let text = '';
        
        // Escolher texto baseado na configuração
        if (config?.karaoke?.text_language === 'translated' && slide.text.translated) {
            text = slide.text.translated;
        } else if (config?.karaoke?.text_language === 'original' || !slide.text.translated) {
            text = slide.text.original;
        } else {
            // "same" - usar o mesmo idioma que o áudio
            text = this.audioLanguage === 'translated' && slide.text.translated
                ? slide.text.translated
                : slide.text.original;
        }
        
        // Dividir em palavras e criar spans
        if (config?.karaoke?.enabled && text) {
            const words = text.split(/\\s+/).filter(w => w.length > 0);
            this.currentWordCount = words.length;  // Guardar para recalcular em onAudioLoaded
            
            this.elements.karaokeText.innerHTML = words.map((word, i) => 
                `<span class="karaoke-word" data-index="${i}">${word}</span>`
            ).join(' ');
            
            // Calcular timings iniciais (serão recalculados quando o áudio carregar)
            this.calculateWordTimings(words.length, slide.duration || 5);
        } else {
            this.currentWordCount = 0;
            this.elements.karaokeText.textContent = text || '';
        }
    }
    
    calculateWordTimings(wordCount, duration) {
        // Timing aproximado por palavra
        this.wordTimings = [];
        const timePerWord = duration / wordCount;
        
        for (let i = 0; i < wordCount; i++) {
            this.wordTimings.push({
                start: i * timePerWord,
                end: (i + 1) * timePerWord
            });
        }
        
        this.currentWordIndex = 0;
    }
    
    loadAudio(slide) {
        let audioSrc = '';
        
        if (this.audioLanguage === 'translated' && slide.audio.translated) {
            audioSrc = slide.audio.translated;
        } else {
            audioSrc = slide.audio.original;
        }
        
        if (audioSrc) {
            this.audioElement.src = audioSrc;
            this.audioElement.load();
        }
    }
    
    togglePlay() {
        if (this.isPlaying) {
            this.pause();
        } else {
            this.play();
        }
    }
    
    play() {
        if (!this.audioElement.src) return;
        
        this.audioElement.play().catch(e => console.log('Erro ao reproduzir:', e));
        this.isPlaying = true;
        if (this.elements?.playBtn) {
            this.elements.playBtn.textContent = '⏸';
        }
        this.startKaraoke();
    }
    
    pause() {
        this.audioElement.pause();
        this.isPlaying = false;
        if (this.elements?.playBtn) {
            this.elements.playBtn.textContent = '▶';
        }
        this.stopKaraoke();
    }
    
    startKaraoke() {
        if (!this.presentationData?.config?.karaoke?.enabled) return;
        
        this.karaokeInterval = setInterval(() => {
            this.updateKaraoke();
        }, 50);
    }
    
    stopKaraoke() {
        if (this.karaokeInterval) {
            clearInterval(this.karaokeInterval);
            this.karaokeInterval = null;
        }
    }
    
    updateKaraoke() {
        const currentTime = this.audioElement.currentTime;
        
        // Encontrar palavra actual
        for (let i = 0; i < this.wordTimings.length; i++) {
            const timing = this.wordTimings[i];
            
            if (currentTime >= timing.start && currentTime < timing.end) {
                if (i !== this.currentWordIndex) {
                    this.highlightWord(i);
                }
                break;
            }
        }
    }
    
    highlightWord(index) {
        const words = this.elements.karaokeText?.querySelectorAll('.karaoke-word');
        if (!words) return;
        
        words.forEach((word, i) => {
            word.classList.remove('active', 'past');
            if (i < index) {
                word.classList.add('past');
            } else if (i === index) {
                word.classList.add('active');
                
                // Scroll se necessário (modo scroll)
                if (this.presentationData?.config?.karaoke?.scroll_mode === 'scroll') {
                    word.scrollIntoView({ behavior: 'smooth', block: 'center' });
                }
            }
        });
        
        this.currentWordIndex = index;
    }
    
    onTimeUpdate() {
        const current = this.audioElement.currentTime;
        const duration = this.audioElement.duration || 0;
        
        // Actualizar barra de progresso
        if (this.elements.progressFill && duration > 0) {
            const percent = (current / duration) * 100;
            this.elements.progressFill.style.width = `${percent}%`;
        }
        
        // Actualizar tempo
        if (this.elements.currentTime) {
            this.elements.currentTime.textContent = this.formatTime(current);
        }
    }
    
    onAudioLoaded() {
        const realDuration = this.audioElement.duration;
        
        if (this.elements.totalTime) {
            this.elements.totalTime.textContent = this.formatTime(realDuration);
        }
        
        // IMPORTANTE: Recalcular timings do karaoke com a duração REAL do áudio
        // Isto corrige o problema de áudios traduzidos com duração diferente
        if (this.currentWordCount > 0 && realDuration > 0) {
            this.calculateWordTimings(this.currentWordCount, realDuration);
        }
    }
    
    onAudioEnded() {
        this.isPlaying = false;
        this.elements.playBtn.textContent = '▶';
        this.stopKaraoke();
        
        // Auto-avançar
        if (this.presentationData?.config?.audio?.autoplay) {
            setTimeout(() => this.nextSlide(), 500);
        }
    }
    
    seekTo(e) {
        const rect = this.elements.progressBar.getBoundingClientRect();
        const percent = (e.clientX - rect.left) / rect.width;
        this.audioElement.currentTime = percent * this.audioElement.duration;
    }
    
    prevSlide() {
        if (this.currentSlide > 0) {
            this.loadSlide(this.currentSlide - 1);
        }
    }
    
    nextSlide() {
        if (this.currentSlide < this.presentationData.slides.length - 1) {
            this.loadSlide(this.currentSlide + 1);
        }
    }
    
    setLanguage(lang) {
        this.audioLanguage = lang;
        this.updateLanguageButtons();
        
        // Recarregar slide actual com novo idioma
        this.loadSlide(this.currentSlide);
    }
    
    toggleLanguage() {
        this.setLanguage(this.audioLanguage === 'original' ? 'translated' : 'original');
    }
    
    updateLanguageButtons() {
        if (!this.elements) return;
        this.elements.langOriginal?.classList.toggle('active', this.audioLanguage === 'original');
        this.elements.langTranslated?.classList.toggle('active', this.audioLanguage === 'translated');
    }
    
    updateSlideCounter() {
        if (this.elements.slideCounter) {
            const total = this.presentationData?.slides?.length || 0;
            this.elements.slideCounter.textContent = `${this.currentSlide + 1} / ${total}`;
        }
    }
    
    toggleFullscreen() {
        const container = document.querySelector('.player-container');
        
        if (document.fullscreenElement) {
            document.exitFullscreen();
        } else {
            container?.requestFullscreen();
        }
    }
    
    formatTime(seconds) {
        if (isNaN(seconds)) return '0:00';
        const mins = Math.floor(seconds / 60);
        const secs = Math.floor(seconds % 60);
        return `${mins}:${secs.toString().padStart(2, '0')}`;
    }
}

// Iniciar player quando DOM estiver pronto
document.addEventListener('DOMContentLoaded', () => {
    window.player = new PresentationPlayer();
});
'''
        return js
    
    def _generate_html_folder(self, css_path: str, js_path: str, presentation_data: dict) -> str:
        """
        Gera o HTML principal (modo pasta) COM DADOS EMBEDADOS.
        Evita problemas de CORS ao abrir via file://
        """
        layout_class = f"layout-{self.config.layout.mode}"
        
        # Converter dados para JSON inline
        data_json = json.dumps(presentation_data, ensure_ascii=False)
        
        html = f'''<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <title>{self.title}</title>
    <link rel="stylesheet" href="{css_path}">
</head>
<body>
    <div class="player-container {layout_class}">
        <div class="slide-area">
            <div class="slide-image-container">
                <img id="slide-image" class="slide-image" src="" alt="Slide">
                <div class="loading hidden">
                    <div class="loading-spinner"></div>
                </div>
            </div>
            
            <div class="karaoke-area">
                <div id="karaoke-text" class="karaoke-text"></div>
            </div>
        </div>
        
        <div class="controls-area">
            <div class="progress-bar" id="progress-bar">
                <div class="progress-bar-fill" id="progress-fill"></div>
            </div>
            
            <div class="controls-buttons">
                <button class="control-btn" id="prev-btn" title="Anterior (←)">⏮</button>
                <button class="control-btn play-btn" id="play-btn" title="Play/Pause (Espaço)">▶</button>
                <button class="control-btn" id="next-btn" title="Próximo (→)">⏭</button>
                
                <div class="volume-control">
                    <span>🔊</span>
                    <input type="range" class="volume-slider" id="volume-slider" 
                           min="0" max="100" value="80" title="Volume (↑/↓)">
                </div>
                
                <div class="language-selector">
                    <button class="lang-btn active" id="lang-original" title="Idioma original">PT</button>
                    <button class="lang-btn" id="lang-translated" title="Tradução (L)">EN</button>
                </div>
                
                <button class="control-btn" id="fullscreen-btn" title="Ecrã completo (F)">⛶</button>
            </div>
            
            <div class="slide-info">
                <span id="current-time">0:00</span>
                <span id="slide-counter">1 / 1</span>
                <span id="total-time">0:00</span>
            </div>
        </div>
    </div>
    
    <!-- Dados da apresentação embedados (evita CORS) -->
    <script id="presentation-data" type="application/json">
{data_json}
    </script>
    
    <script src="{js_path}"></script>
</body>
</html>'''
        return html
    
    def _generate_html(self, css_path: str, js_path: str, data_path: str) -> str:
        """Gera o HTML principal (modo pasta)"""
        layout_class = f"layout-{self.config.layout.mode}"
        
        html = f'''<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <title>{self.title}</title>
    <link rel="stylesheet" href="{css_path}">
</head>
<body>
    <div class="player-container {layout_class}">
        <div class="slide-area">
            <div class="slide-image-container">
                <img id="slide-image" class="slide-image" src="" alt="Slide">
                <div class="loading hidden">
                    <div class="loading-spinner"></div>
                </div>
            </div>
            
            <div class="karaoke-area">
                <div id="karaoke-text" class="karaoke-text"></div>
            </div>
        </div>
        
        <div class="controls-area">
            <div class="progress-bar" id="progress-bar">
                <div class="progress-bar-fill" id="progress-fill"></div>
            </div>
            
            <div class="controls-buttons">
                <button class="control-btn" id="prev-btn" title="Anterior (←)">⏮</button>
                <button class="control-btn play-btn" id="play-btn" title="Play/Pause (Espaço)">▶</button>
                <button class="control-btn" id="next-btn" title="Próximo (→)">⏭</button>
                
                <div class="volume-control">
                    <span>🔊</span>
                    <input type="range" class="volume-slider" id="volume-slider" 
                           min="0" max="100" value="80" title="Volume (↑/↓)">
                </div>
                
                <div class="language-selector">
                    <button class="lang-btn active" id="lang-original" title="Idioma original">PT</button>
                    <button class="lang-btn" id="lang-translated" title="Tradução (L)">EN</button>
                </div>
                
                <button class="control-btn" id="fullscreen-btn" title="Ecrã completo (F)">⛶</button>
            </div>
            
            <div class="slide-info">
                <span id="current-time">0:00</span>
                <span id="slide-counter">1 / 1</span>
                <span id="total-time">0:00</span>
            </div>
        </div>
    </div>
    
    <script src="{js_path}"></script>
</body>
</html>'''
        return html
    
    def _generate_html_single(self, css: str, js: str, data: dict) -> str:
        """Gera HTML com tudo inline (modo ficheiro único)"""
        layout_class = f"layout-{self.config.layout.mode}"
        data_json = json.dumps(data, ensure_ascii=False)
        
        html = f'''<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=no">
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <title>{self.title}</title>
    <style>
{css}
    </style>
</head>
<body>
    <div class="player-container {layout_class}">
        <div class="slide-area">
            <div class="slide-image-container">
                <img id="slide-image" class="slide-image" src="" alt="Slide">
                <div class="loading hidden">
                    <div class="loading-spinner"></div>
                </div>
            </div>
            
            <div class="karaoke-area">
                <div id="karaoke-text" class="karaoke-text"></div>
            </div>
        </div>
        
        <div class="controls-area">
            <div class="progress-bar" id="progress-bar">
                <div class="progress-bar-fill" id="progress-fill"></div>
            </div>
            
            <div class="controls-buttons">
                <button class="control-btn" id="prev-btn" title="Anterior (←)">⏮</button>
                <button class="control-btn play-btn" id="play-btn" title="Play/Pause (Espaço)">▶</button>
                <button class="control-btn" id="next-btn" title="Próximo (→)">⏭</button>
                
                <div class="volume-control">
                    <span>🔊</span>
                    <input type="range" class="volume-slider" id="volume-slider" 
                           min="0" max="100" value="80" title="Volume (↑/↓)">
                </div>
                
                <div class="language-selector">
                    <button class="lang-btn active" id="lang-original" title="Idioma original">PT</button>
                    <button class="lang-btn" id="lang-translated" title="Tradução (L)">EN</button>
                </div>
                
                <button class="control-btn" id="fullscreen-btn" title="Ecrã completo (F)">⛶</button>
            </div>
            
            <div class="slide-info">
                <span id="current-time">0:00</span>
                <span id="slide-counter">1 / 1</span>
                <span id="total-time">0:00</span>
            </div>
        </div>
    </div>
    
    <script>
    // Dados da apresentação inline
    window.PRESENTATION_DATA = {data_json};
    </script>
    <script>
{js}
    </script>
</body>
</html>'''
        return html


# Função auxiliar para exportar slides como imagens
def _exportar_slides_como_imagens(caminho_pptx: str, pasta_saida: str, 
                                   progress_callback: Callable = None) -> List[str]:
    """
    Exporta slides do PPTX como imagens usando LibreOffice + PyMuPDF.
    
    Args:
        caminho_pptx: Caminho do ficheiro PPTX
        pasta_saida: Pasta onde guardar as imagens
        progress_callback: Callback de progresso opcional
    
    Returns:
        Lista de caminhos das imagens geradas (ordenadas por slide)
    """
    import subprocess
    import platform
    import uuid
    
    def _report(current, total, msg):
        if progress_callback:
            progress_callback(current, total, msg)
        print(f"[HTML Export] {msg}")  # Debug
    
    def _obter_soffice_path() -> str:
        sistema = platform.system()
        if sistema == "Windows":
            program_files = [
                os.environ.get("ProgramFiles", r"C:\Program Files"),
                os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
            ]
            for pf in program_files:
                if pf:
                    path = os.path.join(pf, "LibreOffice", "program", "soffice.exe")
                    if os.path.exists(path):
                        return path
        elif sistema == "Darwin":
            mac_paths = [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "/Applications/LibreOffice 7.app/Contents/MacOS/soffice",
                "/Applications/LibreOffice 24.app/Contents/MacOS/soffice",
            ]
            for path in mac_paths:
                if os.path.exists(path):
                    return path
        else:
            linux_paths = ["/usr/bin/soffice", "/usr/bin/libreoffice"]
            for path in linux_paths:
                if os.path.exists(path):
                    return path
        return "soffice"
    
    imagens = []
    os.makedirs(pasta_saida, exist_ok=True)
    
    # Criar pasta temporária para PDF (dentro da pasta de saída para evitar problemas)
    temp_dir = os.path.join(pasta_saida, f"_temp_pdf_{uuid.uuid4().hex[:8]}")
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        # PASSO 1: Converter PPTX para PDF com LibreOffice
        _report(5, 100, "A converter PPTX para PDF...")
        
        soffice = _obter_soffice_path()
        caminho_pptx = os.path.normpath(os.path.abspath(caminho_pptx))
        
        print(f"[HTML Export] PPTX: {caminho_pptx}")
        print(f"[HTML Export] Pasta saída: {pasta_saida}")
        print(f"[HTML Export] LibreOffice: {soffice}")
        
        cmd = [soffice, "--headless", "--invisible", "--convert-to", "pdf", 
               "--outdir", temp_dir, caminho_pptx]
        
        run_kwargs = {"capture_output": True, "timeout": 180}
        if platform.system() == "Windows":
            run_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        
        result = subprocess.run(cmd, **run_kwargs)
        
        if result.returncode != 0:
            erro = result.stderr.decode(errors='ignore') if result.stderr else "Desconhecido"
            _report(0, 100, f"ERRO: LibreOffice falhou: {erro}")
            return []
        
        # Encontrar PDF gerado
        nome_base = Path(caminho_pptx).stem
        pdf_path = os.path.join(temp_dir, f"{nome_base}.pdf")
        
        if not os.path.exists(pdf_path):
            # Tentar encontrar qualquer PDF na pasta
            for f in os.listdir(temp_dir):
                if f.lower().endswith('.pdf'):
                    pdf_path = os.path.join(temp_dir, f)
                    print(f"[HTML Export] PDF encontrado: {pdf_path}")
                    break
        
        if not os.path.exists(pdf_path):
            _report(0, 100, "ERRO: PDF não foi gerado")
            print(f"[HTML Export] Conteúdo de temp_dir: {os.listdir(temp_dir)}")
            return []
        
        print(f"[HTML Export] PDF gerado: {pdf_path}")
        
        # PASSO 2: Converter PDF para imagens
        _report(20, 100, "A converter PDF para imagens...")
        
        # Método 1: PyMuPDF (recomendado para Windows)
        try:
            import fitz  # PyMuPDF
            print("[HTML Export] A usar PyMuPDF")
            
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            print(f"[HTML Export] PDF tem {total_pages} páginas")
            
            for i, pagina in enumerate(doc):
                mat = fitz.Matrix(150/72, 150/72)  # 150 DPI
                pix = pagina.get_pixmap(matrix=mat)
                
                # Guardar na pasta de saída (não na temp)
                img_path = os.path.join(pasta_saida, f"slide_{i+1:03d}.jpg")
                pix.save(img_path)
                imagens.append(img_path)
                
                print(f"[HTML Export] Imagem criada: {img_path}")
                
                _report(20 + int(60 * (i+1) / total_pages), 100, 
                       f"Slide {i+1}/{total_pages} convertido")
            
            doc.close()
            
        except ImportError:
            print("[HTML Export] PyMuPDF não disponível, a tentar pdf2image...")
            # Método 2: pdf2image (precisa Poppler)
            try:
                from pdf2image import convert_from_path
                
                pdf2img_kwargs = {"dpi": 150}
                if platform.system() == "Windows":
                    poppler_paths = [
                        r"C:\Program Files\poppler-24.08.0\Library\bin",
                        r"C:\Program Files\poppler\Library\bin",
                        r"C:\poppler\Library\bin",
                    ]
                    for pp in poppler_paths:
                        if os.path.exists(os.path.join(pp, "pdftoppm.exe")):
                            pdf2img_kwargs["poppler_path"] = pp
                            break
                
                paginas = convert_from_path(pdf_path, **pdf2img_kwargs)
                
                for i, pagina in enumerate(paginas):
                    img_path = os.path.join(pasta_saida, f"slide_{i+1:03d}.jpg")
                    pagina.save(img_path, "JPEG", quality=90)
                    imagens.append(img_path)
                    print(f"[HTML Export] Imagem criada: {img_path}")
                    
            except Exception as e:
                _report(0, 100, f"ERRO: {e}. Instale PyMuPDF: pip install PyMuPDF")
                print(f"[HTML Export] Exceção: {e}")
                return []
        
        print(f"[HTML Export] Total de imagens geradas: {len(imagens)}")
        
    except Exception as e:
        print(f"[HTML Export] Exceção geral: {e}")
        import traceback
        traceback.print_exc()
        return []
        
    finally:
        # Limpar APENAS a pasta temporária do PDF (não as imagens!)
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except:
            pass
    
    return imagens


# Função auxiliar para exportar a partir do GestorPPTX
def exportar_pptx_para_html(gestor_pptx, output_path: str, config: ConfigExportHTML = None,
                            progress_callback: Callable = None) -> bool:
    """
    Função auxiliar para exportar apresentação para HTML5.
    
    Args:
        gestor_pptx: Instância do GestorPPTX com apresentação carregada
        output_path: Caminho de saída
        config: Configuração de exportação (opcional)
        progress_callback: Callback de progresso (opcional)
    
    Returns:
        True se sucesso
    """
    import tempfile
    
    print(f"[HTML Export] Iniciando exportação para: {output_path}")
    
    exporter = HTMLExporter()
    
    if config:
        exporter.config = config
        print(f"[HTML Export] Formato: {config.format}")
    else:
        print("[HTML Export] Config não fornecida, usando padrão")
    
    if progress_callback:
        exporter.set_progress_callback(progress_callback)
    
    # Obter caminho do PPTX
    caminho_pptx = None
    if gestor_pptx.apresentacao:
        caminho_pptx = gestor_pptx.apresentacao.caminho
    
    if not caminho_pptx or not os.path.exists(caminho_pptx):
        print(f"[HTML Export] ERRO: PPTX não encontrado: {caminho_pptx}")
        if progress_callback:
            progress_callback(0, 100, "ERRO: Ficheiro PPTX não encontrado")
        return False
    
    print(f"[HTML Export] PPTX encontrado: {caminho_pptx}")
    
    # Determinar pasta para imagens temporárias
    # Para modo "folder", usar subpasta dentro do output para evitar problemas de limpeza
    formato = config.format if config else "folder"
    
    if formato == "folder":
        # Criar estrutura de pastas primeiro
        base_path = Path(output_path)
        base_path.mkdir(parents=True, exist_ok=True)
        temp_img_dir = str(base_path / "_temp_slides")
    else:
        temp_img_dir = tempfile.mkdtemp(prefix="html_export_slides_")
    
    os.makedirs(temp_img_dir, exist_ok=True)
    print(f"[HTML Export] Pasta temp imagens: {temp_img_dir}")
    
    try:
        # EXPORTAR IMAGENS DOS SLIDES
        if progress_callback:
            progress_callback(0, 100, "A exportar imagens dos slides...")
        
        imagens_slides = _exportar_slides_como_imagens(
            caminho_pptx, 
            temp_img_dir,
            progress_callback
        )
        
        if not imagens_slides:
            print("[HTML Export] ERRO: Nenhuma imagem exportada")
            if progress_callback:
                progress_callback(0, 100, "ERRO: Não foi possível exportar imagens dos slides")
            return False
        
        print(f"[HTML Export] {len(imagens_slides)} imagens exportadas")
        
        # Verificar que as imagens existem
        for img in imagens_slides:
            if os.path.exists(img):
                print(f"[HTML Export] OK: {img}")
            else:
                print(f"[HTML Export] FALTA: {img}")
        
        # Converter slides do GestorPPTX para SlideDataHTML
        slides_data = []
        num_slides = gestor_pptx.num_slides
        
        print(f"[HTML Export] Processando {num_slides} slides...")
        
        for i in range(1, num_slides + 1):
            slide_info = gestor_pptx.obter_slide(i)
            
            # Obter caminho da imagem (índice i-1 porque lista começa em 0)
            img_path = ""
            if i <= len(imagens_slides):
                img_path = imagens_slides[i - 1]
                print(f"[HTML Export] Slide {i} -> imagem: {img_path}")
            else:
                print(f"[HTML Export] Slide {i} -> SEM IMAGEM!")
            
            if slide_info:
                slide_data = SlideDataHTML(
                    id=i,
                    image_path=img_path,
                    audio_original=slide_info.caminho_audio or "",
                    audio_translated=slide_info.caminho_audio_traduzido or "",
                    duration=slide_info.duracao_audio or 5.0,
                    text_original=slide_info.texto_narrar or "",
                    text_translated=slide_info.texto_traduzido or ""
                )
            else:
                slide_data = SlideDataHTML(
                    id=i,
                    image_path=img_path,
                    duration=5.0
                )
            
            slides_data.append(slide_data)
        
        # Título da apresentação
        title = Path(caminho_pptx).stem if caminho_pptx else "Apresentação"
        
        print(f"[HTML Export] A chamar exporter.export()...")
        
        # Exportar HTML
        result = exporter.export(output_path, slides_data, title)
        
        print(f"[HTML Export] Resultado: {result}")
        
        return result
        
    finally:
        # Limpar pasta temporária de imagens APENAS no modo "single"
        # No modo "folder", as imagens são copiadas pelo exporter
        if formato == "single":
            try:
                print(f"[HTML Export] A limpar temp: {temp_img_dir}")
                shutil.rmtree(temp_img_dir, ignore_errors=True)
            except:
                pass
        elif formato == "folder":
            # Limpar a pasta _temp_slides após o export ter copiado as imagens
            try:
                print(f"[HTML Export] A limpar temp slides: {temp_img_dir}")
                shutil.rmtree(temp_img_dir, ignore_errors=True)
            except:
                pass
