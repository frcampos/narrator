#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tradutor.py - Sistema de traduÃ§Ã£o de texto
Suporta Google Translate (online) e Argos Translate (offline)
"""

import os
from typing import Optional
from dataclasses import dataclass

# Google Translate
try:
    from deep_translator import GoogleTranslator
    GOOGLE_DISPONIVEL = True
except ImportError:
    GOOGLE_DISPONIVEL = False

# Argos Translate (offline)
try:
    import argostranslate.package
    import argostranslate.translate
    ARGOS_DISPONIVEL = True
except ImportError:
    ARGOS_DISPONIVEL = False


# Mapeamento de cÃ³digos de idioma para traduÃ§Ã£o
CODIGOS_IDIOMA = {
    "pt-PT": "pt",
    "pt-BR": "pt", 
    "en": "en",
    "en-GB": "en",
    "es": "es",
    "fr": "fr",
    "de": "de",
    "it": "it",
    # Novos idiomas v1.8
    "nl-NL": "nl",
    "pl-PL": "pl",
    "ro-RO": "ro",
    "uk-UA": "uk",
    "el-GR": "el",
    # Línguas do Sul da Ásia v2.0
    "ur-PK": "ur",
    "bn-BD": "bn",
    "bn-IN": "bn",
    "hi-IN": "hi",
    "ta-IN": "ta",
    "te-IN": "te",
    "mr-IN": "mr",
    "gu-IN": "gu",
    "kn-IN": "kn",
    "ml-IN": "ml",
    "pa-IN": "pa",
}

NOMES_IDIOMA = {
    "pt-PT": "Português (Portugal)",
    "pt-BR": "Português (Brasil)",
    "en": "English (US)",
    "en-GB": "English (UK)",
    "es": "Espanhol",
    "fr": "Francês",
    "de": "Alemão",
    "it": "Italiano",
    # Novos idiomas v1.8
    "nl-NL": "Paises Baixos",
    "pl-PL": "Polaco",
    "ro-RO": "Romeno",
    "uk-UA": "Ucrâniano",
    "el-GR": "Grego",
    # Línguas do Sul da Ásia v2.0
    "ur-PK": "Urdu (Pakistan)",
    "bn-BD": "Bengali (Bangladesh)",
    "bn-IN": "Bengali (India)",
    "hi-IN": "Hindi (India)",
    "ta-IN": "Tamil (India)",
    "te-IN": "Telugu (India)",
    "mr-IN": "Marathi (India)",
    "gu-IN": "Gujarati (India)",
    "kn-IN": "Kannada (India)",
    "ml-IN": "Malayalam (India)",
    "pa-IN": "Punjabi (India)",
}

# Limite de caracteres por requisição (margem de segurança)
LIMITE_CHARS_TRADUCAO = 4000


@dataclass
class ConfigTradutor:
    """ConfiguraÃ§Ã£o do tradutor"""
    motor: str = "google"  # "google" ou "argos"
    idioma_origem: str = "pt-PT"
    idioma_destino: str = "en"
    ativo: bool = False


class Tradutor:
    """Classe para traduÃ§Ã£o de texto"""
    
    def __init__(self):
        self.config = ConfigTradutor()
        self._pacotes_argos_instalados = []
    
    @staticmethod
    def google_disponivel() -> bool:
        return GOOGLE_DISPONIVEL
    
    @staticmethod
    def argos_disponivel() -> bool:
        return ARGOS_DISPONIVEL
    
    def motor_disponivel(self) -> bool:
        """Verifica se o motor configurado estÃ¡ disponÃ­vel"""
        if self.config.motor == "google":
            return GOOGLE_DISPONIVEL
        elif self.config.motor == "argos":
            return ARGOS_DISPONIVEL
        return False
    
    def _obter_codigo(self, idioma: str) -> str:
        """Converte cÃ³digo de idioma para formato simples"""
        return CODIGOS_IDIOMA.get(idioma, idioma[:2])
    
    def _instalar_pacote_argos(self, origem: str, destino: str) -> bool:
        """Instala pacote de traduÃ§Ã£o Argos se necessÃ¡rio"""
        if not ARGOS_DISPONIVEL:
            return False
        
        try:
            # Verificar se jÃ¡ estÃ¡ instalado
            installed = argostranslate.translate.get_installed_languages()
            origem_lang = None
            destino_lang = None
            
            for lang in installed:
                if lang.code == origem:
                    origem_lang = lang
                if lang.code == destino:
                    destino_lang = lang
            
            if origem_lang and destino_lang:
                # Verificar se a traduÃ§Ã£o existe
                translation = origem_lang.get_translation(destino_lang)
                if translation:
                    return True
            
            # Descarregar e instalar pacote
            argostranslate.package.update_package_index()
            available = argostranslate.package.get_available_packages()
            
            for pkg in available:
                if pkg.from_code == origem and pkg.to_code == destino:
                    argostranslate.package.install_from_path(pkg.download())
                    return True
            
            return False
            
        except Exception as e:
            print(f"Erro ao instalar pacote Argos: {e}")
            return False
    
    def traduzir(self, texto: str) -> Optional[str]:
        """
        Traduz texto usando o motor configurado.
        
        Returns:
            Texto traduzido ou None se falhar
        """
        if not texto or not texto.strip():
            return ""
        
        if not self.config.ativo:
            return texto
        
        origem = self._obter_codigo(self.config.idioma_origem)
        destino = self._obter_codigo(self.config.idioma_destino)
        
        if origem == destino:
            return texto
        
        try:
            if self.config.motor == "google":
                return self._traduzir_google(texto, origem, destino)
            elif self.config.motor == "argos":
                return self._traduzir_argos(texto, origem, destino)
            else:
                return None
        except Exception as e:
            print(f"Erro na traduÃ§Ã£o: {e}")
            return None
    
    def _traduzir_google(self, texto: str, origem: str, destino: str) -> Optional[str]:
        """
        Traduz usando Google Translate.
        v1.8: Suporte a textos longos com chunking inteligente.
        """
        if not GOOGLE_DISPONIVEL:
            return None
        
        try:
            # v1.8: Se texto Ã© longo, dividir em chunks
            if len(texto) > LIMITE_CHARS_TRADUCAO:
                return self._traduzir_google_chunked(texto, origem, destino)
            
            tradutor = GoogleTranslator(source=origem, target=destino)
            return tradutor.translate(texto)
        except Exception as e:
            print(f"Erro Google Translate: {e}")
            return None
    
    def _traduzir_google_chunked(self, texto: str, origem: str, destino: str) -> Optional[str]:
        """
        Traduz texto longo dividindo em chunks por pontuaÃ§Ã£o.
        v1.8: Nova funÃ§Ã£o para textos > 4000 caracteres.
        """
        import re
        
        chunks = self._dividir_texto_chunks(texto, LIMITE_CHARS_TRADUCAO)
        
        if not chunks:
            return None
        
        tradutor = GoogleTranslator(source=origem, target=destino)
        resultados = []
        
        for i, chunk in enumerate(chunks):
            try:
                resultado = tradutor.translate(chunk)
                if resultado:
                    resultados.append(resultado)
                else:
                    # Se falhar um chunk, manter original
                    resultados.append(chunk)
            except Exception as e:
                print(f"Erro ao traduzir chunk {i+1}/{len(chunks)}: {e}")
                resultados.append(chunk)  # Manter original em caso de erro
        
        return ' '.join(resultados)
    
    def _dividir_texto_chunks(self, texto: str, limite: int) -> list:
        """
        Divide texto em chunks respeitando pontuaÃ§Ã£o.
        v1.8: Chunking inteligente por frases.
        
        Prioridade de divisÃ£o:
        1. Por parÃ¡grafo (\\n\\n)
        2. Por frase (. ! ?)
        3. Por vÃ­rgula/ponto-vÃ­rgula
        4. Por espaÃ§o (Ãºltimo recurso)
        """
        import re
        
        texto = texto.strip()
        if len(texto) <= limite:
            return [texto]
        
        chunks = []
        
        # Tentar dividir por parÃ¡grafos primeiro
        paragrafos = texto.split('\n\n')
        
        chunk_actual = ""
        for paragrafo in paragrafos:
            paragrafo = paragrafo.strip()
            if not paragrafo:
                continue
            
            # Se o parÃ¡grafo cabe no chunk actual
            if len(chunk_actual) + len(paragrafo) + 2 <= limite:
                chunk_actual = (chunk_actual + "\n\n" + paragrafo).strip()
            else:
                # Guardar chunk actual se nÃ£o vazio
                if chunk_actual:
                    chunks.append(chunk_actual)
                
                # Se parÃ¡grafo Ã© maior que limite, dividir por frases
                if len(paragrafo) > limite:
                    sub_chunks = self._dividir_por_frases(paragrafo, limite)
                    chunks.extend(sub_chunks[:-1])  # Adicionar todos menos o Ãºltimo
                    chunk_actual = sub_chunks[-1] if sub_chunks else ""
                else:
                    chunk_actual = paragrafo
        
        # Adicionar Ãºltimo chunk
        if chunk_actual:
            chunks.append(chunk_actual)
        
        return chunks if chunks else [texto[:limite]]
    
    def _dividir_por_frases(self, texto: str, limite: int) -> list:
        """Divide texto por frases (pontuaÃ§Ã£o final)."""
        import re
        
        # Dividir por pontuaÃ§Ã£o final de frase
        frases = re.split(r'(?<=[.!?])\s+', texto)
        
        chunks = []
        chunk_actual = ""
        
        for frase in frases:
            frase = frase.strip()
            if not frase:
                continue
            
            if len(chunk_actual) + len(frase) + 1 <= limite:
                chunk_actual = (chunk_actual + " " + frase).strip()
            else:
                if chunk_actual:
                    chunks.append(chunk_actual)
                
                # Se frase Ã© maior que limite, dividir por vÃ­rgula/espaÃ§o
                if len(frase) > limite:
                    sub_chunks = self._dividir_por_virgula(frase, limite)
                    chunks.extend(sub_chunks[:-1])
                    chunk_actual = sub_chunks[-1] if sub_chunks else ""
                else:
                    chunk_actual = frase
        
        if chunk_actual:
            chunks.append(chunk_actual)
        
        return chunks if chunks else [texto[:limite]]
    
    def _dividir_por_virgula(self, texto: str, limite: int) -> list:
        """Divide texto por vÃ­rgulas ou espaÃ§os (Ãºltimo recurso)."""
        import re
        
        # Tentar dividir por vÃ­rgula/ponto-vÃ­rgula
        partes = re.split(r'(?<=[,;])\s*', texto)
        
        chunks = []
        chunk_actual = ""
        
        for parte in partes:
            parte = parte.strip()
            if not parte:
                continue
            
            if len(chunk_actual) + len(parte) + 1 <= limite:
                chunk_actual = (chunk_actual + " " + parte).strip()
            else:
                if chunk_actual:
                    chunks.append(chunk_actual)
                
                # Se ainda Ã© maior, cortar por espaÃ§o
                if len(parte) > limite:
                    palavras = parte.split()
                    chunk_actual = ""
                    for palavra in palavras:
                        if len(chunk_actual) + len(palavra) + 1 <= limite:
                            chunk_actual = (chunk_actual + " " + palavra).strip()
                        else:
                            if chunk_actual:
                                chunks.append(chunk_actual)
                            chunk_actual = palavra
                else:
                    chunk_actual = parte
        
        if chunk_actual:
            chunks.append(chunk_actual)
        
        return chunks if chunks else [texto[:limite]]
    
    def _traduzir_argos(self, texto: str, origem: str, destino: str) -> Optional[str]:
        """Traduz usando Argos Translate (offline)"""
        if not ARGOS_DISPONIVEL:
            return None
        
        try:
            # Garantir que o pacote estÃ¡ instalado
            self._instalar_pacote_argos(origem, destino)
            
            # Obter idiomas instalados
            installed = argostranslate.translate.get_installed_languages()
            origem_lang = None
            destino_lang = None
            
            for lang in installed:
                if lang.code == origem:
                    origem_lang = lang
                if lang.code == destino:
                    destino_lang = lang
            
            if not origem_lang or not destino_lang:
                return None
            
            translation = origem_lang.get_translation(destino_lang)
            if translation:
                return translation.translate(texto)
            
            return None
            
        except Exception as e:
            print(f"Erro Argos Translate: {e}")
            return None
    
    def traduzir_lote(self, textos: list) -> list:
        """Traduz lista de textos"""
        return [self.traduzir(t) for t in textos]


# InstÃ¢ncia global
tradutor = Tradutor()
