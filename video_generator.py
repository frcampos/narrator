#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
video_generator.py - Geração de vídeo a partir de slides
v1.8 - Correções Windows: caminhos temporários, gestão de ficheiros
"""

import os
import subprocess
import tempfile
import uuid
import shutil
import atexit
from pathlib import Path
from typing import List, Optional, Callable
from dataclasses import dataclass

# MoviePy
try:
    from moviepy import AudioFileClip, ImageClip, concatenate_videoclips
    MOVIEPY_V2 = True
except ImportError:
    try:
        from moviepy.audio.io.AudioFileClip import AudioFileClip
        from moviepy.video.VideoClip import ImageClip
        from moviepy.video.compositing.concatenate import concatenate_videoclips
        MOVIEPY_V2 = False
    except ImportError:
        MOVIEPY_V2 = None

from PIL import Image

from pptx_handler import GestorPPTX, SlideInfo


@dataclass
class ConfigVideo:
    """Configuração de vídeo"""
    largura: int = 1920
    altura: int = 1080
    fps: int = 24
    tempo_minimo_slide: float = 3.0
    tempo_extra_slide: float = 0.5
    duracao_transicao: float = 0.5
    codec_video: str = "libx264"
    codec_audio: str = "aac"
    # Legendas
    legendas_embutidas: bool = False
    legendas_usar_traducao: bool = True
    legendas_tamanho_fonte: int = 28
    legendas_cor: str = "white"
    legendas_fundo: bool = True
    legendas_posicao: str = "sobrepor"
    legendas_linhas: int = 3
    # Karaoke
    karaoke_ativo: bool = False
    karaoke_modo: str = "scroll"
    karaoke_cor: str = "Yellow"
    karaoke_transparencia: int = 70
    karaoke_usar_traducao: bool = False


CORES_WEB_SAFE = {
    "Yellow": (255, 255, 0),
    "Cyan": (0, 255, 255),
    "Lime": (0, 255, 0),
    "Magenta": (255, 0, 255),
    "Orange": (255, 165, 0),
    "Pink": (255, 192, 203),
    "Aqua": (0, 255, 255),
    "Red": (255, 0, 0),
    "Green": (0, 128, 0),
    "Blue": (0, 0, 255),
    "White": (255, 255, 255),
}


class GeradorVideo:
    """Gerador de vídeo a partir de apresentações"""
    
    def __init__(self):
        self.config = ConfigVideo()
        self._progresso_callback: Optional[Callable[[int, int, str], None]] = None
        self._pasta_temp_base: Optional[str] = None
        self._ficheiros_temp: List[str] = []
        # Registar limpeza ao sair
        atexit.register(self._limpar_todos_temp)
    
    def _criar_pasta_temp(self) -> str:
        """Cria pasta temporária segura para a sessão"""
        if self._pasta_temp_base and os.path.exists(self._pasta_temp_base):
            return self._pasta_temp_base
        
        # Usar pasta temp do sistema com UUID único
        base = tempfile.gettempdir()
        pasta = os.path.join(base, f"pptx_narrator_{uuid.uuid4().hex[:8]}")
        os.makedirs(pasta, exist_ok=True)
        self._pasta_temp_base = pasta
        return pasta
    
    def _criar_ficheiro_temp(self, sufixo: str = ".jpg") -> str:
        """Cria caminho para ficheiro temporário de forma segura (Windows-compatible)"""
        pasta = self._criar_pasta_temp()
        nome = f"temp_{uuid.uuid4().hex[:12]}{sufixo}"
        caminho = os.path.join(pasta, nome)
        self._ficheiros_temp.append(caminho)
        return caminho
    
    def _limpar_todos_temp(self):
        """Limpa todos os ficheiros temporários"""
        for f in self._ficheiros_temp:
            try:
                if os.path.exists(f):
                    os.remove(f)
            except:
                pass
        self._ficheiros_temp.clear()
        
        if self._pasta_temp_base and os.path.exists(self._pasta_temp_base):
            try:
                shutil.rmtree(self._pasta_temp_base, ignore_errors=True)
            except:
                pass
            self._pasta_temp_base = None
    
    @staticmethod
    def disponivel() -> bool:
        """Verifica se MoviePy está disponível"""
        return MOVIEPY_V2 is not None
    
    @staticmethod
    def _obter_soffice_path() -> str:
        """Obtém caminho do executável soffice"""
        import platform
        
        sistema = platform.system()
        
        if sistema == "Windows":
            # Procurar em Program Files
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
                os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            ]
            for path in mac_paths:
                if os.path.exists(path):
                    return path
        
        else:
            linux_paths = [
                "/usr/bin/soffice",
                "/usr/bin/libreoffice",
                "/usr/local/bin/soffice",
            ]
            for path in linux_paths:
                if os.path.exists(path):
                    return path
        
        return "soffice"
    
    @staticmethod
    def libreoffice_disponivel() -> bool:
        """Verifica se LibreOffice está disponível"""
        soffice = GeradorVideo._obter_soffice_path()
        
        if os.path.isabs(soffice):
            return os.path.exists(soffice)
        
        try:
            # Windows: usar shell=True para encontrar no PATH
            import platform
            if platform.system() == "Windows":
                result = subprocess.run(
                    f'"{soffice}" --version',
                    capture_output=True, 
                    timeout=10,
                    shell=True
                )
            else:
                result = subprocess.run(
                    [soffice, "--version"], 
                    capture_output=True, 
                    timeout=10
                )
            return result.returncode == 0
        except:
            return False
    
    def definir_callback_progresso(self, callback: Callable[[int, int, str], None]):
        """Define callback para progresso: (atual, total, mensagem)"""
        self._progresso_callback = callback
    
    def _reportar_progresso(self, atual: int, total: int, msg: str):
        """Reporta progresso"""
        if self._progresso_callback:
            self._progresso_callback(atual, total, msg)
    
    def _executar_comando(self, cmd: List[str], timeout: int = 180) -> subprocess.CompletedProcess:
        """Executa comando de forma segura no Windows e outros sistemas"""
        import platform
        
        kwargs = {
            "capture_output": True,
            "timeout": timeout,
        }
        
        # Windows: configurações específicas para evitar janelas popup
        if platform.system() == "Windows":
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
            # Normalizar caminhos para Windows
            cmd = [str(c).replace("/", "\\") if os.path.exists(str(c)) else str(c) for c in cmd]
        
        return subprocess.run(cmd, **kwargs)
    
    def _exportar_slides_imagens(self, caminho_pptx: str, pasta_saida: str) -> List[str]:
        """Exporta slides como imagens usando LibreOffice."""
        if not self.libreoffice_disponivel():
            self._reportar_progresso(0, 100, "ERRO: LibreOffice não encontrado!")
            print("LibreOffice não encontrado. Instale em: https://www.libreoffice.org/")
            return []
        
        os.makedirs(pasta_saida, exist_ok=True)
        soffice = self._obter_soffice_path()
        
        # Normalizar caminhos para o sistema
        caminho_pptx = os.path.normpath(os.path.abspath(caminho_pptx))
        pasta_saida = os.path.normpath(os.path.abspath(pasta_saida))
        
        try:
            self._reportar_progresso(5, 100, "A converter PPTX para PDF...")
            
            cmd_pdf = [
                soffice,
                "--headless",
                "--invisible",
                "--convert-to", "pdf",
                "--outdir", pasta_saida,
                caminho_pptx
            ]
            
            result = self._executar_comando(cmd_pdf)
            
            if result.returncode != 0:
                erro = result.stderr.decode(errors='ignore') if result.stderr else "Erro desconhecido"
                self._reportar_progresso(0, 100, f"Erro LibreOffice: {erro}")
                print(f"Erro LibreOffice: {erro}")
                return []
            
            # Encontrar PDF gerado
            nome_base = Path(caminho_pptx).stem
            pdf_path = os.path.join(pasta_saida, f"{nome_base}.pdf")
            
            if not os.path.exists(pdf_path):
                for f in os.listdir(pasta_saida):
                    if f.lower().endswith('.pdf'):
                        pdf_path = os.path.join(pasta_saida, f)
                        break
            
            if not os.path.exists(pdf_path):
                self._reportar_progresso(0, 100, "ERRO: PDF não foi gerado")
                return []
            
            self._reportar_progresso(15, 100, "A converter PDF para imagens...")
            
            imagens = self._pdf_para_imagens(pdf_path, pasta_saida)
            
            return imagens
            
        except subprocess.TimeoutExpired:
            print("Timeout ao converter com LibreOffice")
            return []
        except Exception as e:
            print(f"Erro ao exportar slides: {e}")
            return []
    
    def _pdf_para_imagens(self, pdf_path: str, pasta_saida: str) -> List[str]:
        """Converte PDF para imagens - múltiplos métodos para compatibilidade"""
        imagens = []
        import platform
        sistema = platform.system()
        
        # MÉTODO 1: PyMuPDF (fitz) - RECOMENDADO para Windows
        # Funciona sem dependências externas
        try:
            import fitz  # PyMuPDF
            self._reportar_progresso(16, 100, "A usar PyMuPDF...")
            
            doc = fitz.open(pdf_path)
            for i, pagina in enumerate(doc):
                # Renderizar a 150 DPI (matriz de escala)
                mat = fitz.Matrix(150/72, 150/72)
                pix = pagina.get_pixmap(matrix=mat)
                
                caminho = os.path.join(pasta_saida, f"slide_{i+1:03d}.jpg")
                pix.save(caminho)
                imagens.append(caminho)
                
                self._reportar_progresso(
                    16 + int(4 * (i+1) / len(doc)), 100,
                    f"Slide {i+1}/{len(doc)} convertido"
                )
            
            doc.close()
            
            if imagens:
                return imagens
        except ImportError:
            print("PyMuPDF não instalado. Instale com: pip install PyMuPDF")
        except Exception as e:
            print(f"Erro PyMuPDF: {e}")
        
        # MÉTODO 2: pdf2image com Poppler (precisa Poppler instalado)
        try:
            from pdf2image import convert_from_path
            self._reportar_progresso(16, 100, "A usar pdf2image...")
            
            kwargs = {"dpi": 150}
            
            if sistema == "Windows":
                # Procurar Poppler em vários locais comuns
                poppler_paths = [
                    r"C:\Program Files\poppler-24.08.0\Library\bin",
                    r"C:\Program Files\poppler\Library\bin",
                    r"C:\Program Files\poppler\bin",
                    r"C:\poppler-24.08.0\Library\bin",
                    r"C:\poppler\Library\bin",
                    r"C:\poppler\bin",
                    os.path.join(os.environ.get("LOCALAPPDATA", ""), "poppler", "Library", "bin"),
                    os.path.join(os.environ.get("PROGRAMFILES", ""), "poppler", "Library", "bin"),
                ]
                
                poppler_encontrado = False
                for pp in poppler_paths:
                    pdftoppm_path = os.path.join(pp, "pdftoppm.exe")
                    if os.path.exists(pdftoppm_path):
                        kwargs["poppler_path"] = pp
                        poppler_encontrado = True
                        print(f"Poppler encontrado em: {pp}")
                        break
                
                if not poppler_encontrado:
                    print("AVISO: Poppler não encontrado. pdf2image pode falhar.")
                    print("Instale PyMuPDF (pip install PyMuPDF) como alternativa.")
            
            paginas = convert_from_path(pdf_path, **kwargs)
            
            for i, pagina in enumerate(paginas):
                caminho = os.path.join(pasta_saida, f"slide_{i+1:03d}.jpg")
                pagina.save(caminho, "JPEG", quality=90)
                imagens.append(caminho)
            
            if imagens:
                return imagens
                
        except ImportError:
            print("pdf2image não instalado")
        except Exception as e:
            print(f"Erro pdf2image: {e}")
        
        # MÉTODO 3: pdftoppm direto (Linux/Mac apenas)
        if sistema != "Windows":
            try:
                prefixo = os.path.join(pasta_saida, "slide")
                cmd = ["pdftoppm", "-jpeg", "-r", "150", pdf_path, prefixo]
                result = self._executar_comando(cmd, timeout=120)
                
                if result.returncode == 0:
                    for f in sorted(os.listdir(pasta_saida)):
                        if f.startswith("slide") and f.lower().endswith(".jpg"):
                            imagens.append(os.path.join(pasta_saida, f))
                    
                    if imagens:
                        return imagens
            except Exception as e:
                print(f"Erro pdftoppm: {e}")
        
        # Se chegou aqui, nenhum método funcionou
        if not imagens:
            self._reportar_progresso(0, 100, "ERRO: Instale PyMuPDF (pip install PyMuPDF)")
            print("\n" + "="*60)
            print("ERRO: Não foi possível converter PDF para imagens.")
            print("SOLUÇÃO: Instale PyMuPDF com o comando:")
            print("    pip install PyMuPDF")
            print("="*60 + "\n")
        
        return imagens
    
    def _redimensionar_imagem(self, caminho: str) -> str:
        """Redimensiona imagem para o tamanho do vídeo mantendo proporções"""
        with Image.open(caminho) as img:
            if img.mode in ('RGBA', 'P'):
                img = img.convert('RGB')
            
            img_ratio = img.width / img.height
            video_ratio = self.config.largura / self.config.altura
            
            if img_ratio > video_ratio:
                novo_w = self.config.largura
                novo_h = int(self.config.largura / img_ratio)
            else:
                novo_h = self.config.altura
                novo_w = int(self.config.altura * img_ratio)
            
            img_resized = img.resize((novo_w, novo_h), Image.Resampling.LANCZOS)
            
            resultado = Image.new('RGB', (self.config.largura, self.config.altura), (0, 0, 0))
            pos_x = (self.config.largura - novo_w) // 2
            pos_y = (self.config.altura - novo_h) // 2
            resultado.paste(img_resized, (pos_x, pos_y))
            
            # Usar método seguro para ficheiros temporários
            temp_path = self._criar_ficheiro_temp('.jpg')
            resultado.save(temp_path, 'JPEG', quality=95)
            
            return temp_path
    
    def _obter_fonte(self, tamanho: int):
        """Obtém fonte disponível no sistema."""
        from PIL import ImageFont
        import platform
        
        sistema = platform.system()
        
        fontes_possiveis = []
        
        if sistema == "Windows":
            # Fontes Windows
            windows_fonts = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
            fontes_possiveis = [
                os.path.join(windows_fonts, "arial.ttf"),
                os.path.join(windows_fonts, "calibri.ttf"),
                os.path.join(windows_fonts, "segoeui.ttf"),
                os.path.join(windows_fonts, "tahoma.ttf"),
            ]
        elif sistema == "Darwin":
            fontes_possiveis = [
                "/System/Library/Fonts/Helvetica.ttc",
                "/System/Library/Fonts/SFNSText.ttf",
            ]
        else:
            fontes_possiveis = [
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "/usr/share/fonts/TTF/DejaVuSans.ttf",
            ]
        
        for fonte_path in fontes_possiveis:
            try:
                if os.path.exists(fonte_path):
                    return ImageFont.truetype(fonte_path, tamanho)
            except:
                continue
        
        # Último recurso
        try:
            return ImageFont.truetype("arial.ttf", tamanho)
        except:
            pass
        
        return ImageFont.load_default()
    
    def _quebrar_texto_largura(self, texto: str, fonte, largura_max: int, draw) -> List[str]:
        """Quebra texto em linhas que cabem na largura especificada."""
        palavras = texto.split()
        linhas = []
        linha_atual = []
        
        for palavra in palavras:
            linha_teste = ' '.join(linha_atual + [palavra])
            bbox = draw.textbbox((0, 0), linha_teste, font=fonte)
            largura_linha = bbox[2] - bbox[0]
            
            if largura_linha <= largura_max:
                linha_atual.append(palavra)
            else:
                if linha_atual:
                    linhas.append(' '.join(linha_atual))
                linha_atual = [palavra]
        
        if linha_atual:
            linhas.append(' '.join(linha_atual))
        
        return linhas if linhas else [texto[:50]]
    
    def _dividir_texto_segmentos(self, texto: str) -> List[str]:
        """Divide texto longo em segmentos para legendas dinâmicas."""
        chars_por_linha = 40
        max_chars_segmento = self.config.legendas_linhas * chars_por_linha
        max_chars_segmento = min(max_chars_segmento, 120)
        
        texto = texto.strip().replace('\n', ' ').replace('  ', ' ')
        
        if len(texto) <= max_chars_segmento:
            return [texto]
        
        import re
        frases = re.split(r'(?<=[.!?])\s+', texto)
        
        segmentos = []
        segmento_atual = ""
        
        for frase in frases:
            frase = frase.strip()
            if not frase:
                continue
            
            if len(segmento_atual) + len(frase) + 1 > max_chars_segmento:
                if segmento_atual:
                    segmentos.append(segmento_atual.strip())
                
                if len(frase) > max_chars_segmento:
                    palavras = frase.split()
                    segmento_atual = ""
                    for palavra in palavras:
                        if len(segmento_atual) + len(palavra) + 1 > max_chars_segmento:
                            if segmento_atual:
                                segmentos.append(segmento_atual.strip())
                            segmento_atual = palavra
                        else:
                            segmento_atual = (segmento_atual + " " + palavra).strip()
                else:
                    segmento_atual = frase
            else:
                segmento_atual = (segmento_atual + " " + frase).strip()
        
        if segmento_atual:
            segmentos.append(segmento_atual.strip())
        
        if not segmentos:
            segmentos = [texto[:max_chars_segmento]]
        
        return segmentos
    
    def _adicionar_legenda_imagem(self, caminho_img: str, texto: str) -> str:
        """Adiciona legendas na imagem."""
        from PIL import ImageDraw
        
        modo_sobrepor = self.config.legendas_posicao == "sobrepor"
        max_linhas = self.config.legendas_linhas
        
        with Image.open(caminho_img) as img_original:
            largura_orig, altura_orig = img_original.size
            
            tamanho_fonte = max(20, int(altura_orig * 0.028))
            fonte = self._obter_fonte(tamanho_fonte)
            
            altura_linha = tamanho_fonte + 8
            altura_barra = (max_linhas * altura_linha) + 20
            
            if modo_sobrepor:
                img = img_original.convert('RGBA')
                largura, altura = img.size
                
                overlay = Image.new('RGBA', (largura, altura), (0, 0, 0, 0))
                draw_overlay = ImageDraw.Draw(overlay)
                
                y_barra = altura - altura_barra
                draw_overlay.rectangle(
                    [(0, y_barra), (largura, altura)],
                    fill=(0, 0, 0, 200)
                )
                
                img = Image.alpha_composite(img, overlay)
                draw = ImageDraw.Draw(img)
            else:
                img_original = img_original.convert('RGB')
                nova_altura = altura_orig + altura_barra
                img = Image.new('RGB', (largura_orig, nova_altura), (0, 0, 0))
                img.paste(img_original, (0, 0))
                
                largura, altura = img.size
                y_barra = altura_orig
                draw = ImageDraw.Draw(img)
            
            texto = texto.strip().replace('\n', ' ').replace('  ', ' ')
            
            margem = 40
            largura_max = largura - (margem * 2)
            
            linhas = self._quebrar_texto_largura(texto, fonte, largura_max, draw)
            
            if len(linhas) > max_linhas:
                linhas = linhas[:max_linhas]
                if linhas[-1] and not linhas[-1].endswith('...'):
                    linhas[-1] = linhas[-1][:-3] + '...' if len(linhas[-1]) > 3 else '...'
            
            altura_total_texto = len(linhas) * altura_linha
            y_inicio = y_barra + (altura_barra - altura_total_texto) // 2
            
            for i, linha in enumerate(linhas):
                if not linha:
                    continue
                
                bbox = draw.textbbox((0, 0), linha, font=fonte)
                largura_texto = bbox[2] - bbox[0]
                x = (largura - largura_texto) // 2
                y = y_inicio + i * altura_linha
                
                draw.text((x + 2, y + 2), linha, font=fonte, fill=(0, 0, 0, 255) if modo_sobrepor else (50, 50, 50))
                draw.text((x, y), linha, font=fonte, fill=(255, 255, 255, 255) if modo_sobrepor else (255, 255, 255))
            
            if modo_sobrepor:
                img = img.convert('RGB')
            
            temp_path = self._criar_ficheiro_temp('.jpg')
            img.save(temp_path, 'JPEG', quality=95)
            
            return temp_path
    
    def _gerar_clips_karaoke(self, img_base: str, texto: str, duracao_slide: float, 
                             audio_path: str, ficheiros_temp: List[str],
                             duracao_audio: float = None) -> List:
        """Gera clips com efeito karaoke (destaque palavra-a-palavra)."""
        palavras = texto.strip().split()
        num_palavras = len(palavras)
        
        if num_palavras == 0:
            return []
        
        duracao_para_palavras = duracao_audio if duracao_audio and duracao_audio > 0 else duracao_slide
        
        tempo_por_palavra = duracao_para_palavras / num_palavras
        
        min_tempo_frame = 0.3
        if tempo_por_palavra < min_tempo_frame:
            palavras_por_frame = max(1, int(min_tempo_frame / tempo_por_palavra))
            grupos_palavras = []
            for i in range(0, num_palavras, palavras_por_frame):
                grupos_palavras.append(palavras[i:i + palavras_por_frame])
        else:
            grupos_palavras = [[p] for p in palavras]
        
        num_frames = len(grupos_palavras)
        tempo_por_frame_audio = duracao_para_palavras / num_frames
        tempo_extra_final = max(0, duracao_slide - duracao_para_palavras)
        
        clips = []
        modo_scroll = self.config.karaoke_modo == "scroll"
        
        # Carregar áudio uma vez só
        audio_clip = None
        if audio_path and os.path.exists(audio_path):
            try:
                audio_clip = AudioFileClip(audio_path)
            except Exception as e:
                print(f"Erro ao carregar áudio: {e}")
        
        for idx, grupo in enumerate(grupos_palavras):
            idx_palavra_inicio = sum(len(g) for g in grupos_palavras[:idx])
            idx_palavra_fim = idx_palavra_inicio + len(grupo) - 1
            
            img_karaoke = self._criar_frame_karaoke(
                img_base, palavras, idx_palavra_inicio, idx_palavra_fim, modo_scroll
            )
            ficheiros_temp.append(img_karaoke)
            
            if idx == num_frames - 1:
                duracao_frame = tempo_por_frame_audio + tempo_extra_final
            else:
                duracao_frame = tempo_por_frame_audio
            
            if MOVIEPY_V2:
                clip = ImageClip(img_karaoke, duration=duracao_frame)
            else:
                clip = ImageClip(img_karaoke).set_duration(duracao_frame)
            
            # Áudio apenas no primeiro frame
            if idx == 0 and audio_clip:
                if MOVIEPY_V2:
                    clip = clip.with_audio(audio_clip)
                else:
                    clip = clip.set_audio(audio_clip)
            
            clips.append(clip)
        
        return clips
    
    def _criar_frame_karaoke(self, img_base: str, palavras: List[str], 
                              idx_inicio: int, idx_fim: int, modo_scroll: bool) -> str:
        """Cria um frame de karaoke com palavras destacadas."""
        from PIL import ImageDraw
        
        modo_sobrepor = self.config.legendas_posicao == "sobrepor"
        max_linhas = self.config.legendas_linhas
        
        with Image.open(img_base) as img_original:
            largura_orig, altura_orig = img_original.size
            
            tamanho_fonte = max(20, int(altura_orig * 0.028))
            fonte = self._obter_fonte(tamanho_fonte)
            
            altura_linha = tamanho_fonte + 8
            altura_barra = (max_linhas * altura_linha) + 20
            
            if modo_sobrepor:
                img = img_original.convert('RGBA')
                largura, altura = img.size
                
                overlay = Image.new('RGBA', (largura, altura), (0, 0, 0, 0))
                draw_overlay = ImageDraw.Draw(overlay)
                
                y_barra = altura - altura_barra
                draw_overlay.rectangle(
                    [(0, y_barra), (largura, altura)],
                    fill=(0, 0, 0, 200)
                )
                
                img = Image.alpha_composite(img, overlay)
            else:
                img_original = img_original.convert('RGB')
                nova_altura = altura_orig + altura_barra
                img = Image.new('RGB', (largura_orig, nova_altura), (0, 0, 0))
                img.paste(img_original, (0, 0))
                
                largura, altura = img.size
                y_barra = altura_orig
            
            draw = ImageDraw.Draw(img)
            
            margem = 40
            largura_max = largura - (margem * 2)
            
            linhas_palavras = self._organizar_palavras_linhas(palavras, fonte, largura_max, draw)
            
            if modo_scroll:
                linha_atual = 0
                contador = 0
                for i, linha in enumerate(linhas_palavras):
                    if contador + len(linha) > idx_inicio:
                        linha_atual = i
                        break
                    contador += len(linha)
                
                linhas_antes = max_linhas // 2
                linha_inicio = max(0, linha_atual - linhas_antes)
                linha_fim = min(len(linhas_palavras), linha_inicio + max_linhas)
                
                if linha_fim == len(linhas_palavras):
                    linha_inicio = max(0, linha_fim - max_linhas)
                
                linhas_visiveis = linhas_palavras[linha_inicio:linha_fim]
                idx_offset = sum(len(linhas_palavras[i]) for i in range(linha_inicio))
            else:
                pagina_atual = 0
                contador = 0
                
                for i in range(0, len(linhas_palavras), max_linhas):
                    linhas_pagina = linhas_palavras[i:i + max_linhas]
                    palavras_pagina = sum(len(l) for l in linhas_pagina)
                    if contador + palavras_pagina > idx_inicio:
                        linhas_visiveis = linhas_pagina
                        idx_offset = contador
                        break
                    contador += palavras_pagina
                else:
                    linhas_visiveis = linhas_palavras[:max_linhas]
                    idx_offset = 0
            
            altura_total_texto = len(linhas_visiveis) * altura_linha
            y_inicio = y_barra + (altura_barra - altura_total_texto) // 2
            
            cor_destaque = CORES_WEB_SAFE.get(self.config.karaoke_cor, (255, 255, 0))
            alpha_destaque = int(self.config.karaoke_transparencia * 2.55)
            
            idx_palavra_global = idx_offset
            
            for i, linha in enumerate(linhas_visiveis):
                texto_linha = ' '.join(linha)
                bbox = draw.textbbox((0, 0), texto_linha, font=fonte)
                largura_linha = bbox[2] - bbox[0]
                x = (largura - largura_linha) // 2
                y = y_inicio + i * altura_linha
                
                x_atual = x
                for palavra in linha:
                    bbox_palavra = draw.textbbox((0, 0), palavra, font=fonte)
                    largura_palavra = bbox_palavra[2] - bbox_palavra[0]
                    
                    if idx_inicio <= idx_palavra_global <= idx_fim:
                        padding = 4
                        if modo_sobrepor:
                            destaque_overlay = Image.new('RGBA', img.size, (0, 0, 0, 0))
                            draw_destaque = ImageDraw.Draw(destaque_overlay)
                            draw_destaque.rectangle(
                                [(x_atual - padding, y - padding),
                                 (x_atual + largura_palavra + padding, y + altura_linha - padding)],
                                fill=(*cor_destaque, alpha_destaque)
                            )
                            img = Image.alpha_composite(img, destaque_overlay)
                            draw = ImageDraw.Draw(img)
                        else:
                            cor_fundo_mista = tuple(
                                int(c * (self.config.karaoke_transparencia / 100))
                                for c in cor_destaque
                            )
                            draw.rectangle(
                                [(x_atual - padding, y - padding),
                                 (x_atual + largura_palavra + padding, y + altura_linha - padding)],
                                fill=cor_fundo_mista
                            )
                    
                    cor_texto = (255, 255, 255, 255) if modo_sobrepor else (255, 255, 255)
                    cor_sombra = (0, 0, 0, 255) if modo_sobrepor else (0, 0, 0)
                    
                    draw.text((x_atual + 1, y + 1), palavra, font=fonte, fill=cor_sombra)
                    draw.text((x_atual, y), palavra, font=fonte, fill=cor_texto)
                    
                    espaco = draw.textbbox((0, 0), ' ', font=fonte)[2]
                    x_atual += largura_palavra + espaco
                    idx_palavra_global += 1
            
            if modo_sobrepor:
                img = img.convert('RGB')
            
            temp_path = self._criar_ficheiro_temp('.jpg')
            img.save(temp_path, 'JPEG', quality=90)
            
            return temp_path
    
    def _organizar_palavras_linhas(self, palavras: List[str], fonte, 
                                    largura_max: int, draw) -> List[List[str]]:
        """Organiza palavras em linhas respeitando largura máxima."""
        linhas = []
        linha_atual = []
        largura_atual = 0
        espaco = draw.textbbox((0, 0), ' ', font=fonte)[2]
        
        for palavra in palavras:
            bbox = draw.textbbox((0, 0), palavra, font=fonte)
            largura_palavra = bbox[2] - bbox[0]
            
            if largura_atual + largura_palavra + espaco <= largura_max or not linha_atual:
                linha_atual.append(palavra)
                largura_atual += largura_palavra + espaco
            else:
                linhas.append(linha_atual)
                linha_atual = [palavra]
                largura_atual = largura_palavra + espaco
        
        if linha_atual:
            linhas.append(linha_atual)
        
        return linhas
    
    def gerar_video(
        self,
        gestor_pptx: GestorPPTX,
        caminho_saida: str,
        usar_audio_traduzido: bool = False
    ) -> bool:
        """
        Gera vídeo MP4 a partir da apresentação.
        
        Args:
            gestor_pptx: Gestor com apresentação carregada
            caminho_saida: Caminho para o ficheiro MP4
            usar_audio_traduzido: Se True, usa áudio traduzido em vez do original
            
        Returns:
            True se gerado com sucesso
        """
        if not self.disponivel():
            self._reportar_progresso(0, 100, "ERRO: MoviePy não disponível")
            return False
        
        if not gestor_pptx.apresentacao:
            self._reportar_progresso(0, 100, "ERRO: Apresentação não carregada")
            return False
        
        self._usar_audio_traduzido = usar_audio_traduzido
        
        # Normalizar caminho de saída
        caminho_saida = os.path.normpath(os.path.abspath(caminho_saida))
        
        # Criar pasta temporária para slides
        pasta_temp = self._criar_pasta_temp()
        pasta_slides = os.path.join(pasta_temp, "slides")
        os.makedirs(pasta_slides, exist_ok=True)
        
        self._reportar_progresso(0, 100, "A iniciar exportação de slides...")
        
        imagens = self._exportar_slides_imagens(
            gestor_pptx.apresentacao.caminho, 
            pasta_slides
        )
        
        if not imagens:
            self._reportar_progresso(0, 100, "Erro: não foi possível exportar slides")
            return False
        
        if len(imagens) != gestor_pptx.num_slides:
            print(f"Aviso: {len(imagens)} imagens para {gestor_pptx.num_slides} slides")
        
        try:
            clips = []
            ficheiros_temp = []
            total = len(imagens)
            
            for i, img_path in enumerate(imagens):
                slide_num = i + 1
                
                slide_info = gestor_pptx.obter_slide(slide_num)
                
                duracao_audio = 0
                audio_path = None
                
                if slide_info:
                    caminho_audio_orig = slide_info.caminho_audio
                    caminho_audio_trad = slide_info.caminho_audio_traduzido
                    
                    duracao_orig = 0
                    duracao_trad = 0
                    
                    if caminho_audio_orig and os.path.exists(caminho_audio_orig):
                        try:
                            clip_temp = AudioFileClip(caminho_audio_orig)
                            duracao_orig = clip_temp.duration
                            clip_temp.close()
                        except:
                            duracao_orig = slide_info.duracao_audio or 0
                    
                    if caminho_audio_trad and os.path.exists(caminho_audio_trad):
                        try:
                            clip_temp = AudioFileClip(caminho_audio_trad)
                            duracao_trad = clip_temp.duration
                            clip_temp.close()
                        except:
                            duracao_trad = slide_info.duracao_audio_traduzido or 0
                    
                    duracao_audio_real = 0
                    
                    if getattr(self, '_usar_audio_traduzido', False):
                        if caminho_audio_trad and os.path.exists(caminho_audio_trad):
                            audio_path = caminho_audio_trad
                            duracao_audio_real = duracao_trad
                            self._reportar_progresso(
                                20 + int(50 * i / total), 100,
                                f"Slide {slide_num}: áudio TRADUZIDO ({duracao_trad:.1f}s)"
                            )
                        elif caminho_audio_orig and os.path.exists(caminho_audio_orig):
                            audio_path = caminho_audio_orig
                            duracao_audio_real = duracao_orig
                            self._reportar_progresso(
                                20 + int(50 * i / total), 100,
                                f"Slide {slide_num}: fallback ORIGINAL ({duracao_orig:.1f}s)"
                            )
                    else:
                        if caminho_audio_orig and os.path.exists(caminho_audio_orig):
                            audio_path = caminho_audio_orig
                            duracao_audio_real = duracao_orig
                            self._reportar_progresso(
                                20 + int(50 * i / total), 100,
                                f"Slide {slide_num}: áudio ORIGINAL ({duracao_orig:.1f}s)"
                            )
                    
                    duracao_audio = max(duracao_orig, duracao_trad)
                    
                    if duracao_audio_real <= 0:
                        duracao_audio_real = duracao_audio
                    
                    if duracao_audio <= 0:
                        duracao_audio = max(
                            slide_info.duracao_audio or 0,
                            slide_info.duracao_audio_traduzido or 0
                        )
                else:
                    self._reportar_progresso(
                        20 + int(50 * i / total), 100,
                        f"Slide {slide_num}: sem informação de slide"
                    )
                
                if duracao_audio > 0:
                    duracao = duracao_audio + self.config.tempo_extra_slide
                else:
                    duracao = self.config.tempo_minimo_slide
                
                img_processada = self._redimensionar_imagem(img_path)
                ficheiros_temp.append(img_processada)
                
                if self.config.legendas_embutidas and slide_info:
                    if self.config.karaoke_ativo:
                        if self.config.karaoke_usar_traducao:
                            texto_karaoke = slide_info.texto_traduzido or slide_info.texto_narrar
                        else:
                            texto_karaoke = slide_info.texto_narrar
                        
                        if texto_karaoke and texto_karaoke.strip():
                            self._reportar_progresso(
                                20 + int(50 * i / total), 100,
                                f"Slide {slide_num}: gerando karaoke"
                            )
                            
                            clips_karaoke = self._gerar_clips_karaoke(
                                img_processada, texto_karaoke, duracao, audio_path, ficheiros_temp,
                                duracao_audio_real
                            )
                            
                            if clips_karaoke:
                                clips.extend(clips_karaoke)
                                continue
                    
                    else:
                        texto_legenda = slide_info.texto_traduzido if self.config.legendas_usar_traducao else slide_info.texto_narrar
                        if texto_legenda and texto_legenda.strip():
                            segmentos = self._dividir_texto_segmentos(texto_legenda)
                            num_segmentos = len(segmentos)
                            
                            if num_segmentos > 1:
                                duracao_por_segmento = duracao / num_segmentos
                                
                                for seg_idx, segmento in enumerate(segmentos):
                                    img_com_legenda = self._adicionar_legenda_imagem(
                                        img_processada, segmento
                                    )
                                    ficheiros_temp.append(img_com_legenda)
                                    
                                    if MOVIEPY_V2:
                                        clip_seg = ImageClip(img_com_legenda, duration=duracao_por_segmento)
                                    else:
                                        clip_seg = ImageClip(img_com_legenda).set_duration(duracao_por_segmento)
                                    
                                    if seg_idx == 0 and audio_path:
                                        audio = AudioFileClip(audio_path)
                                        if MOVIEPY_V2:
                                            clip_seg = clip_seg.with_audio(audio)
                                        else:
                                            clip_seg = clip_seg.set_audio(audio)
                                    
                                    clips.append(clip_seg)
                                
                                continue
                            else:
                                img_com_legenda = self._adicionar_legenda_imagem(
                                    img_processada, texto_legenda
                                )
                                ficheiros_temp.append(img_com_legenda)
                                img_processada = img_com_legenda
                
                if MOVIEPY_V2:
                    clip = ImageClip(img_processada, duration=duracao)
                else:
                    clip = ImageClip(img_processada).set_duration(duracao)
                
                if audio_path and os.path.exists(audio_path):
                    try:
                        audio = AudioFileClip(audio_path)
                        if MOVIEPY_V2:
                            clip = clip.with_audio(audio)
                        else:
                            clip = clip.set_audio(audio)
                    except Exception as e:
                        print(f"Erro ao adicionar áudio ao slide {slide_num}: {e}")
                
                clips.append(clip)
            
            if not clips:
                self._reportar_progresso(0, 100, "ERRO: Nenhum clip gerado")
                return False
            
            self._reportar_progresso(75, 100, "A concatenar clips...")
            video_final = concatenate_videoclips(clips, method="compose")
            
            self._reportar_progresso(80, 100, "A exportar vídeo (pode demorar)...")
            
            # Garantir que pasta de destino existe
            pasta_destino = os.path.dirname(caminho_saida)
            if pasta_destino:
                os.makedirs(pasta_destino, exist_ok=True)
            
            video_final.write_videofile(
                caminho_saida,
                codec=self.config.codec_video,
                audio_codec=self.config.codec_audio,
                fps=self.config.fps,
                preset="ultrafast",
                logger=None
            )
            
            # Fechar clips para libertar recursos
            try:
                video_final.close()
            except:
                pass
            
            for clip in clips:
                try:
                    clip.close()
                except:
                    pass
            
            self._reportar_progresso(95, 100, "A limpar ficheiros temporários...")
            
            # Limpar ficheiros temporários desta sessão
            for f in ficheiros_temp:
                if f not in self._ficheiros_temp:
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except:
                        pass
            
            self._limpar_todos_temp()
            
            self._reportar_progresso(100, 100, "Vídeo criado com sucesso!")
            return True
            
        except Exception as e:
            import traceback
            erro_completo = traceback.format_exc()
            print(f"Erro ao gerar vídeo: {e}")
            print(erro_completo)
            self._reportar_progresso(0, 100, f"Erro: {e}")
            return False


# Instância global
gerador_video = GeradorVideo()
