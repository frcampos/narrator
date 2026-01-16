#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tts_manager.py
==============
Gestão de Text-to-Speech com suporte a múltiplos motores.
Suporta: Edge TTS, pyttsx3, gTTS
"""

import os
import asyncio
import logging
from typing import Optional

from pydub import AudioSegment

from config_manager import GestorConfig, ConfiguracaoTTS


class GestorTTS:
    """
    Gestor de Text-to-Speech com suporte a múltiplos motores.
    
    Motores suportados:
    - edge: Microsoft Edge TTS (online, alta qualidade)
    - pyttsx3: Offline, usa vozes do sistema
    - gtts: Google Text-to-Speech (online)
    """
    
    def __init__(self, config: GestorConfig):
        self.config = config
        self.tts_config = config.tts
        self._motor_atual = None
        self._inicializar_motor()
    
    def _inicializar_motor(self):
        """Inicializa o motor TTS configurado ou usa fallback"""
        for motor in self.tts_config.ordem_fallback:
            if self._testar_motor(motor):
                self._motor_atual = motor
                logging.info(f"🔊 Motor TTS inicializado: {motor}")
                return
        
        logging.error("❌ Nenhum motor TTS disponível!")
        self._motor_atual = None
    
    def _testar_motor(self, motor: str) -> bool:
        """Testa se um motor TTS está disponível"""
        try:
            if motor == "edge":
                import edge_tts
                return True
            elif motor == "pyttsx3":
                import pyttsx3
                engine = pyttsx3.init()
                engine.stop()
                return True
            elif motor == "gtts":
                from gtts import gTTS
                return True
        except ImportError:
            logging.debug(f"Motor {motor} não instalado")
            return False
        except Exception as e:
            logging.warning(f"Motor {motor} falhou no teste: {e}")
            return False
        return False
    
    @property
    def motor_disponivel(self) -> bool:
        """Verifica se há um motor TTS disponível"""
        return self._motor_atual is not None
    
    @property
    def motor_atual(self) -> Optional[str]:
        """Retorna o nome do motor atual"""
        return self._motor_atual
    
    # =========================================================================
    # EDGE TTS
    # =========================================================================
    
    def _formatar_velocidade_edge(self) -> str:
        """Formata velocidade para Edge TTS (formato: +X% ou -X%)"""
        vel = self.tts_config.velocidade
        if vel == 1.0:
            return "+0%"
        elif vel > 1.0:
            return f"+{int((vel - 1) * 100)}%"
        else:
            return f"-{int((1 - vel) * 100)}%"
    
    async def _gerar_audio_edge_async(self, texto: str, caminho_saida: str) -> bool:
        """Gera áudio usando Edge TTS (async)"""
        try:
            import edge_tts
            
            communicate = edge_tts.Communicate(
                texto,
                self.tts_config.voz,
                rate=self._formatar_velocidade_edge(),
                pitch=self.tts_config.pitch,
                volume=self.tts_config.volume
            )
            
            await communicate.save(caminho_saida)
            logging.debug(f"Edge TTS: áudio gerado em {caminho_saida}")
            return True
            
        except Exception as e:
            logging.error(f"Erro Edge TTS: {e}")
            return False
    
    def gerar_audio_edge(self, texto: str, caminho_saida: str) -> bool:
        """Gera áudio usando Edge TTS (sync wrapper)"""
        return asyncio.run(self._gerar_audio_edge_async(texto, caminho_saida))
    
    # =========================================================================
    # PYTTSX3
    # =========================================================================
    
    def gerar_audio_pyttsx3(self, texto: str, caminho_saida: str) -> bool:
        """Gera áudio usando pyttsx3 (offline)"""
        try:
            import pyttsx3
            
            engine = pyttsx3.init()
            
            # Tentar encontrar voz no idioma configurado
            voices = engine.getProperty('voices')
            voz_encontrada = False
            
            idioma_lower = self.tts_config.idioma.lower()
            
            for voice in voices:
                voice_id_lower = voice.id.lower()
                # Procurar por português
                if idioma_lower in voice_id_lower or 'portuguese' in voice_id_lower:
                    engine.setProperty('voice', voice.id)
                    voz_encontrada = True
                    logging.debug(f"pyttsx3: usando voz {voice.id}")
                    break
            
            if not voz_encontrada and voices:
                engine.setProperty('voice', voices[0].id)
                logging.debug(f"pyttsx3: usando voz padrão {voices[0].id}")
            
            # Configurar velocidade
            rate = engine.getProperty('rate')
            nova_rate = int(rate * self.tts_config.velocidade)
            engine.setProperty('rate', nova_rate)
            
            # Gerar áudio
            engine.save_to_file(texto, caminho_saida)
            engine.runAndWait()
            engine.stop()
            
            sucesso = os.path.exists(caminho_saida) and os.path.getsize(caminho_saida) > 0
            if sucesso:
                logging.debug(f"pyttsx3: áudio gerado em {caminho_saida}")
            return sucesso
            
        except Exception as e:
            logging.error(f"Erro pyttsx3: {e}")
            return False
    
    # =========================================================================
    # GTTS
    # =========================================================================
    
    def gerar_audio_gtts(self, texto: str, caminho_saida: str) -> bool:
        """Gera áudio usando gTTS (Google Text-to-Speech)"""
        try:
            from gtts import gTTS
            
            # gTTS usa código de idioma simples (pt, en, etc.)
            idioma = self.tts_config.idioma.split('-')[0]
            
            # Velocidade lenta se configurado
            slow = self.tts_config.velocidade < 0.8
            
            tts = gTTS(text=texto, lang=idioma, slow=slow)
            tts.save(caminho_saida)
            
            logging.debug(f"gTTS: áudio gerado em {caminho_saida}")
            return True
            
        except Exception as e:
            logging.error(f"Erro gTTS: {e}")
            return False
    
    # =========================================================================
    # MÉTODO PRINCIPAL
    # =========================================================================
    
    def gerar_audio(self, texto: str, caminho_saida: str) -> bool:
        """
        Gera áudio a partir de texto usando o motor disponível.
        
        Args:
            texto: Texto a converter em áudio
            caminho_saida: Caminho para guardar o ficheiro de áudio
            
        Returns:
            True se o áudio foi gerado com sucesso
        """
        if not texto.strip():
            logging.warning("Texto vazio, áudio não gerado")
            return False
        
        if not self._motor_atual:
            logging.error("Nenhum motor TTS disponível")
            return False
        
        # Garantir que a pasta existe
        pasta = os.path.dirname(caminho_saida)
        if pasta:
            os.makedirs(pasta, exist_ok=True)
        
        # Determinar extensão e caminho temporário
        extensao_final = os.path.splitext(caminho_saida)[1].lower()
        
        if extensao_final == '.mp3':
            caminho_temp = caminho_saida
        else:
            caminho_temp = caminho_saida.rsplit('.', 1)[0] + '.mp3'
        
        # Gerar áudio com o motor atual
        sucesso = False
        
        if self._motor_atual == "edge":
            sucesso = self.gerar_audio_edge(texto, caminho_temp)
        elif self._motor_atual == "pyttsx3":
            sucesso = self.gerar_audio_pyttsx3(texto, caminho_temp)
        elif self._motor_atual == "gtts":
            sucesso = self.gerar_audio_gtts(texto, caminho_temp)
        
        # Converter formato se necessário
        if sucesso and caminho_temp != caminho_saida:
            try:
                if extensao_final == '.wav':
                    audio = AudioSegment.from_mp3(caminho_temp)
                    audio.export(caminho_saida, format='wav')
                    os.remove(caminho_temp)
                else:
                    os.rename(caminho_temp, caminho_saida)
            except Exception as e:
                logging.error(f"Erro ao converter formato: {e}")
                return False
        
        return sucesso
    
    def obter_duracao_audio(self, caminho: str) -> float:
        """
        Obtém a duração de um ficheiro de áudio em segundos.
        
        Args:
            caminho: Caminho para o ficheiro de áudio
            
        Returns:
            Duração em segundos (0.0 se erro)
        """
        try:
            audio = AudioSegment.from_file(caminho)
            return len(audio) / 1000.0
        except Exception as e:
            logging.error(f"Erro ao obter duração de {caminho}: {e}")
            return 0.0


# =============================================================================
# FUNÇÕES AUXILIARES
# =============================================================================

def listar_vozes_disponiveis():
    """Lista todas as vozes TTS disponíveis no sistema"""
    print("\n" + "="*70)
    print("🔊 VOZES TTS DISPONÍVEIS")
    print("="*70)
    
    # Edge TTS - Português Europeu
    print("\n📢 EDGE TTS - Português Europeu (pt-PT):")
    print("   • pt-PT-RaquelNeural   - Feminina (muito natural, RECOMENDADA)")
    print("   • pt-PT-FernandaNeural - Feminina")
    print("   • pt-PT-DuarteNeural   - Masculino (muito natural, RECOMENDADO)")
    
    # Edge TTS - Português Brasileiro
    print("\n📢 EDGE TTS - Português Brasileiro (pt-BR):")
    print("   • pt-BR-FranciscaNeural - Feminina (muito natural)")
    print("   • pt-BR-AntonioNeural   - Masculino")
    print("   • pt-BR-ThalitaNeural   - Feminina")
    print("   • pt-BR-LeticiaNeural   - Feminina")
    print("   • pt-BR-ManuelaNeural   - Feminina")
    print("   • pt-BR-NicolauNeural   - Masculino")
    print("   • pt-BR-ValerioNeural   - Masculino")
    print("   • pt-BR-YaraNeural      - Feminina")
    
    # pyttsx3 - Vozes do sistema
    print("\n📢 PYTTSX3 - Vozes do Sistema:")
    try:
        import pyttsx3
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        
        if voices:
            for voice in voices[:10]:  # Limitar a 10
                nome = voice.name if hasattr(voice, 'name') else voice.id
                print(f"   • {nome}")
            if len(voices) > 10:
                print(f"   ... e mais {len(voices) - 10} vozes")
        else:
            print("   (nenhuma voz encontrada)")
        
        engine.stop()
    except ImportError:
        print("   (pyttsx3 não instalado)")
    except Exception as e:
        print(f"   (erro ao listar: {e})")
    
    # gTTS
    print("\n📢 GTTS - Google Text-to-Speech:")
    print("   • Suporta múltiplos idiomas via código (pt, en, es, fr, etc.)")
    print("   • Requer ligação à internet")
    
    print("\n" + "="*70)
    print("💡 DICA: Configure a voz no config.ini na secção [TTS]")
    print("="*70 + "\n")


async def testar_edge_tts():
    """Testa se Edge TTS está funcional"""
    try:
        import edge_tts
        
        voices = await edge_tts.list_voices()
        pt_voices = [v for v in voices if v['Locale'].startswith('pt-')]
        
        print("\n🔊 Vozes Edge TTS disponíveis para Português:")
        for v in pt_voices:
            genero = "♀️" if v['Gender'] == 'Female' else "♂️"
            print(f"   {genero} {v['ShortName']} - {v['Locale']}")
        
        return True
    except Exception as e:
        print(f"❌ Erro ao testar Edge TTS: {e}")
        return False


def testar_tts():
    """Testa todos os motores TTS disponíveis"""
    print("\n🧪 TESTE DE MOTORES TTS")
    print("="*50)
    
    # Testar Edge TTS
    print("\n1. Edge TTS...")
    try:
        import edge_tts
        print("   ✅ Disponível")
        asyncio.run(testar_edge_tts())
    except ImportError:
        print("   ❌ Não instalado (pip install edge-tts)")
    
    # Testar pyttsx3
    print("\n2. pyttsx3...")
    try:
        import pyttsx3
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        engine.stop()
        print(f"   ✅ Disponível ({len(voices)} vozes)")
    except ImportError:
        print("   ❌ Não instalado (pip install pyttsx3)")
    except Exception as e:
        print(f"   ⚠️ Erro: {e}")
    
    # Testar gTTS
    print("\n3. gTTS...")
    try:
        from gtts import gTTS
        print("   ✅ Disponível")
    except ImportError:
        print("   ❌ Não instalado (pip install gtts)")
    
    print("\n" + "="*50)


if __name__ == "__main__":
    # Teste rápido quando executado diretamente
    testar_tts()
    listar_vozes_disponiveis()
