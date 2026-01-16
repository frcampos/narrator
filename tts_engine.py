#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tts_engine.py - Motor de Text-to-Speech
Suporta:
- Edge TTS (online, alta qualidade)
- Piper TTS (offline, qualidade neural)
- pyttsx3 (offline, qualidade básica)
"""

import os
import asyncio
import tempfile
import subprocess
from typing import Optional, List, Tuple
from dataclasses import dataclass


# Verificar disponibilidade dos motores
def _verificar_edge():
    try:
        import edge_tts
        return True
    except ImportError:
        return False

def _verificar_piper():
    try:
        result = subprocess.run(['piper', '--help'], 
                                capture_output=True, timeout=5)
        return True
    except:
        return False

def _verificar_pyttsx3():
    try:
        import pyttsx3
        return True
    except ImportError:
        return False

EDGE_DISPONIVEL = _verificar_edge()
PIPER_DISPONIVEL = _verificar_piper()
PYTTSX3_DISPONIVEL = _verificar_pyttsx3()

# Reprodução de áudio
try:
    import pygame
    pygame.mixer.init()
    PYGAME_DISPONIVEL = True
except:
    PYGAME_DISPONIVEL = False


# Vozes Piper disponíveis por idioma
VOZES_PIPER = {
    "pt-PT": [
        ("pt_PT-tugao-medium", "Tugão", "Voz masculina portuguesa"),
    ],
    "pt-BR": [
        ("pt_BR-faber-medium", "Faber", "Voz masculina brasileira"),
    ],
    "en": [
        ("en_US-amy-medium", "Amy (US)", "Voz feminina americana"),
        ("en_US-ryan-medium", "Ryan (US)", "Voz masculina americana"),
        ("en_GB-alan-medium", "Alan (UK)", "Voz masculina britânica"),
    ],
    "es": [
        ("es_ES-sharvard-medium", "Sharvard", "Voz masculina espanhola"),
        ("es_MX-ald-medium", "Ald (MX)", "Voz masculina mexicana"),
    ],
    "fr": [
        ("fr_FR-upmc-medium", "UPMC", "Voz masculina francesa"),
        ("fr_FR-siwis-medium", "Siwis", "Voz feminina francesa"),
    ],
}


# Descrições dos motores para tooltips
DESCRICOES_MOTORES = {
    "edge": "Vozes de alta qualidade da Microsoft. Requer ligação à internet.",
    "piper": "Vozes neurais offline. Qualidade próxima do online, sem internet.",
    "pyttsx3": "Vozes do sistema operativo. Funciona offline mas qualidade básica.",
}


@dataclass
class ConfigTTS:
    """Configuração do motor TTS"""
    motor: str = "edge"  # "edge", "piper", "pyttsx3"
    voz: str = "pt-PT-RaquelNeural"
    idioma: str = "pt-PT"
    velocidade: float = 1.0


class MotorTTS:
    """Motor de síntese de voz unificado"""
    
    def __init__(self):
        self.config = ConfigTTS()
        self._engine_pyttsx3 = None
    
    @staticmethod
    def edge_disponivel() -> bool:
        return EDGE_DISPONIVEL
    
    @staticmethod
    def piper_disponivel() -> bool:
        return PIPER_DISPONIVEL
    
    @staticmethod
    def pyttsx3_disponivel() -> bool:
        return PYTTSX3_DISPONIVEL
    
    @staticmethod
    def offline_disponivel() -> bool:
        return PIPER_DISPONIVEL or PYTTSX3_DISPONIVEL
    
    def motor_disponivel(self) -> bool:
        if self.config.motor == "edge":
            return EDGE_DISPONIVEL
        elif self.config.motor == "piper":
            return PIPER_DISPONIVEL
        elif self.config.motor == "pyttsx3":
            return PYTTSX3_DISPONIVEL
        return False
    
    @staticmethod
    def obter_motores_disponiveis() -> List[Tuple[str, str, str]]:
        """Retorna lista de motores disponíveis (código, nome, descrição)"""
        motores = []
        if EDGE_DISPONIVEL:
            motores.append(("edge", "Edge TTS (Online)", DESCRICOES_MOTORES["edge"]))
        if PIPER_DISPONIVEL:
            motores.append(("piper", "Piper TTS (Offline Neural)", DESCRICOES_MOTORES["piper"]))
        if PYTTSX3_DISPONIVEL:
            motores.append(("pyttsx3", "Sistema (Offline Básico)", DESCRICOES_MOTORES["pyttsx3"]))
        return motores
    
    def obter_vozes_disponiveis(self) -> List[Tuple[str, str]]:
        """Retorna lista de vozes para o motor e idioma atual"""
        if self.config.motor == "piper":
            vozes = VOZES_PIPER.get(self.config.idioma, [])
            if not vozes:
                codigo_simples = self.config.idioma.split('-')[0]
                for codigo in VOZES_PIPER:
                    if codigo.startswith(codigo_simples):
                        vozes = VOZES_PIPER[codigo]
                        break
            return [(v[0], v[1]) for v in vozes]
        else:
            from idiomas import VOZES_EDGE
            vozes = VOZES_EDGE.get(self.config.idioma, [])
            # Retornar apenas código e nome (ignorar género)
            return [(v[0], v[1]) for v in vozes]
    
    def gerar_audio(self, texto: str, caminho_saida: str) -> bool:
        if not texto.strip():
            return False
        
        if self.config.motor == "edge":
            return self._gerar_edge(texto, caminho_saida)
        elif self.config.motor == "piper":
            return self._gerar_piper(texto, caminho_saida)
        elif self.config.motor == "pyttsx3":
            return self._gerar_pyttsx3(texto, caminho_saida)
        return False
    
    def _formatar_velocidade_edge(self) -> str:
        vel = self.config.velocidade
        if vel == 1.0:
            return "+0%"
        elif vel > 1.0:
            return f"+{int((vel - 1) * 100)}%"
        else:
            return f"-{int((1 - vel) * 100)}%"
    
    def _gerar_edge(self, texto: str, caminho_saida: str) -> bool:
        if not EDGE_DISPONIVEL:
            return False
        
        try:
            import edge_tts
            
            async def gerar():
                communicate = edge_tts.Communicate(
                    texto,
                    self.config.voz,
                    rate=self._formatar_velocidade_edge()
                )
                await communicate.save(caminho_saida)
            
            try:
                asyncio.run(gerar())
            except RuntimeError:
                loop = asyncio.new_event_loop()
                asyncio.set_event_loop(loop)
                try:
                    loop.run_until_complete(gerar())
                finally:
                    loop.close()
            
            return os.path.exists(caminho_saida)
            
        except Exception as e:
            print(f"Erro Edge TTS: {e}")
            return False
    
    def _gerar_piper(self, texto: str, caminho_saida: str) -> bool:
        if not PIPER_DISPONIVEL:
            return False
        
        try:
            modelo = self.config.voz
            length_scale = 1.0 / self.config.velocidade
            wav_temp = tempfile.mktemp(suffix='.wav')
            
            process = subprocess.Popen(
                ['piper', '--model', modelo, '--output_file', wav_temp,
                 '--length_scale', str(length_scale)],
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            stdout, stderr = process.communicate(input=texto.encode('utf-8'), timeout=120)
            
            if os.path.exists(wav_temp):
                if caminho_saida.endswith('.mp3'):
                    try:
                        from pydub import AudioSegment
                        audio = AudioSegment.from_wav(wav_temp)
                        audio.export(caminho_saida, format="mp3", bitrate="128k")
                        os.remove(wav_temp)
                        return os.path.exists(caminho_saida)
                    except:
                        import shutil
                        shutil.move(wav_temp, caminho_saida.replace('.mp3', '.wav'))
                        return True
                else:
                    import shutil
                    shutil.move(wav_temp, caminho_saida)
                    return True
            
            return False
            
        except Exception as e:
            print(f"Erro Piper TTS: {e}")
            return False
    
    def _gerar_pyttsx3(self, texto: str, caminho_saida: str) -> bool:
        if not PYTTSX3_DISPONIVEL:
            return False
        
        try:
            import pyttsx3
            
            engine = pyttsx3.init()
            voices = engine.getProperty('voices')
            idioma = self.config.idioma.lower()
            
            for voice in voices:
                if idioma in voice.id.lower() or 'portuguese' in voice.id.lower():
                    engine.setProperty('voice', voice.id)
                    break
            
            rate = engine.getProperty('rate')
            engine.setProperty('rate', int(rate * self.config.velocidade))
            
            engine.save_to_file(texto, caminho_saida)
            engine.runAndWait()
            engine.stop()
            
            return os.path.exists(caminho_saida)
            
        except Exception as e:
            print(f"Erro pyttsx3: {e}")
            return False
    
    # Aliases para compatibilidade
    def gerar_edge(self, texto: str, caminho: str) -> bool:
        return self._gerar_edge(texto, caminho)
    
    def gerar_offline(self, texto: str, caminho: str) -> bool:
        if PIPER_DISPONIVEL:
            return self._gerar_piper(texto, caminho)
        return self._gerar_pyttsx3(texto, caminho)
    
    def gerar_preview(self, texto: str) -> Optional[str]:
        caminho = tempfile.mktemp(suffix='.mp3')
        if self.gerar_audio(texto, caminho):
            return caminho
        return None
    
    def tocar_audio(self, caminho: str):
        # Parar qualquer áudio anterior
        self.parar_audio()
        
        if not PYGAME_DISPONIVEL:
            try:
                import platform
                sistema = platform.system()
                if sistema == "Darwin":
                    self._processo_audio = subprocess.Popen(['afplay', caminho])
                elif sistema == "Windows":
                    os.startfile(caminho)
                    self._processo_audio = None
                else:
                    # Linux - usar mpv, ffplay ou aplay
                    for player in ['mpv --no-video', 'ffplay -nodisp -autoexit', 'aplay']:
                        try:
                            cmd = player.split() + [caminho]
                            self._processo_audio = subprocess.Popen(cmd, 
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                            break
                        except:
                            continue
            except Exception as e:
                print(f"Erro ao reproduzir: {e}")
            return
        
        try:
            pygame.mixer.music.load(caminho)
            pygame.mixer.music.play()
        except Exception as e:
            print(f"Erro ao reproduzir: {e}")
    
    def parar_audio(self):
        # Parar processo externo se existir
        if hasattr(self, '_processo_audio') and self._processo_audio:
            try:
                self._processo_audio.terminate()
                self._processo_audio.kill()
                self._processo_audio = None
            except:
                pass
        
        # Parar pygame se disponível
        if PYGAME_DISPONIVEL:
            try:
                pygame.mixer.music.stop()
            except:
                pass
    
    @staticmethod
    def obter_duracao(caminho: str) -> float:
        """
        Obtém duração de um ficheiro de áudio em segundos.
        Compatível com Python 3.13 (não depende de audioop).
        """
        if not caminho or not os.path.exists(caminho):
            return 0.0
        
        # Método 1: Usar mutagen (mais fiável para MP3)
        try:
            from mutagen.mp3 import MP3
            audio = MP3(caminho)
            return audio.info.length
        except:
            pass
        
        # Método 2: Usar ffprobe (ffmpeg)
        try:
            cmd = [
                "ffprobe", "-v", "quiet", "-show_entries",
                "format=duration", "-of", "default=noprint_wrappers=1:nokey=1",
                caminho
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
            if result.returncode == 0 and result.stdout.strip():
                return float(result.stdout.strip())
        except:
            pass
        
        # Método 3: Usar moviepy
        try:
            from moviepy import AudioFileClip
            clip = AudioFileClip(caminho)
            duracao = clip.duration
            clip.close()
            return duracao
        except:
            pass
        
        # Método 4: Tentar pydub (pode falhar em Python 3.13)
        try:
            from pydub import AudioSegment
            audio = AudioSegment.from_file(caminho)
            return len(audio) / 1000.0
        except:
            pass
        
        return 0.0
