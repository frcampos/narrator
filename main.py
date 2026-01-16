#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
main.py - CONVERSOR MULTIMÉDIA EDUCATIVO v3.0
==============================================
Orquestrador principal para todas as funcionalidades.

Uso:
    python main.py audio              - Processa ficheiros de áudio
    python main.py pptx <ficheiro>    - Processa apresentação PPTX
    python main.py vozes              - Lista vozes TTS disponíveis
    python main.py --help             - Mostra ajuda
"""

import os
import sys
import argparse
import logging

from config_manager import GestorConfig, configurar_logging, criar_pastas
from audio_video_converter import ConversorMultimedia, processar_pasta_audio
from tts_manager import GestorTTS, listar_vozes_disponiveis, testar_tts
from pptx_processor import ProcessadorPPTX, ExtratorPPTX, PPTX_DISPONIVEL


def processar_audio(caminho_config: str = "config.ini"):
    """Processa todos os ficheiros de áudio na pasta de entrada."""
    print("\n" + "="*60)
    print("🎵 PROCESSAMENTO DE ÁUDIO")
    print("="*60)
    
    if not os.path.exists(caminho_config):
        print(f"❌ Ficheiro de configuração não encontrado: {caminho_config}")
        return False
    
    config = GestorConfig(caminho_config)
    configurar_logging(config)
    criar_pastas(config)
    
    logging.info("🎵 Iniciando processamento de áudio")
    ficheiros = processar_pasta_audio(config)
    
    print("\n" + "-"*40)
    if ficheiros:
        print(f"✅ {len(ficheiros)} ficheiro(s) gerado(s)")
        for f in ficheiros:
            print(f"   • {f}")
    else:
        print("⚠️ Nenhum ficheiro processado")
    print("-"*40)
    
    return len(ficheiros) > 0


def processar_pptx_comando(
    caminho_pptx: str,
    caminho_config: str = "config.ini",
    ficheiro_edicao: str = None,
    gerar_video: bool = False,
    apenas_extrair: bool = False
):
    """Processa uma apresentação PowerPoint."""
    print("\n" + "="*60)
    print("📊 PROCESSAMENTO DE POWERPOINT")
    print("="*60)
    
    if not PPTX_DISPONIVEL:
        print("❌ python-pptx não está instalado!")
        print("   Instale com: pip install python-pptx")
        return False
    
    if not os.path.exists(caminho_config):
        print(f"❌ Config não encontrado: {caminho_config}")
        return False
    
    if not os.path.exists(caminho_pptx):
        print(f"❌ PPTX não encontrado: {caminho_pptx}")
        return False
    
    config = GestorConfig(caminho_config)
    configurar_logging(config)
    criar_pastas(config)
    
    print(f"📄 Ficheiro: {caminho_pptx}")
    
    if apenas_extrair:
        extrator = ExtratorPPTX(config)
        slides = extrator.extrair_slides(caminho_pptx)
        ficheiro = extrator.guardar_ficheiro_edicao(
            slides, caminho_pptx, config.pastas.pasta_intermedios
        )
        print(f"\n✅ Texto extraído para: {ficheiro}")
        print("   Edite o ficheiro e execute novamente com -e")
        return True
    
    processador = ProcessadorPPTX(config)
    
    slides_texto = None
    if ficheiro_edicao and os.path.exists(ficheiro_edicao):
        print(f"📝 A usar texto editado de: {ficheiro_edicao}")
        slides_texto = processador.extrator.carregar_ficheiro_edicao(ficheiro_edicao)
    
    resultado = processador.processar_pptx(caminho_pptx, slides_texto, gerar_video)
    
    print("\n" + "-"*40)
    print("📋 RESULTADO:")
    print("-"*40)
    
    if resultado.get("erro"):
        print(f"❌ Erro: {resultado['erro']}")
        return False
    
    if resultado.get("ficheiro_edicao"):
        print(f"📝 Ficheiro edição: {resultado['ficheiro_edicao']}")
    if resultado.get("audios_gerados"):
        print(f"🔊 Áudios: {resultado['slides_processados']} slides")
    if resultado.get("pptx_com_audio"):
        print(f"📊 PPTX: {resultado['pptx_com_audio']}")
    if resultado.get("video_gerado"):
        print(f"🎬 Vídeo: {resultado['video_gerado']}")
    
    print("-"*40)
    return True


def mostrar_ajuda():
    """Mostra ajuda de utilização."""
    ajuda = """
╔══════════════════════════════════════════════════════════════════╗
║           CONVERSOR MULTIMÉDIA EDUCATIVO v3.0                    ║
╠══════════════════════════════════════════════════════════════════╣
║                                                                  ║
║  COMANDOS:                                                       ║
║  ─────────                                                       ║
║                                                                  ║
║  audio                 Processa áudio para MP3/MP4               ║
║    python main.py audio                                          ║
║    python main.py audio -c config.ini                            ║
║                                                                  ║
║  pptx <ficheiro>       Processa apresentação PowerPoint          ║
║    python main.py pptx aula.pptx                                 ║
║    python main.py pptx aula.pptx --extrair                       ║
║    python main.py pptx aula.pptx -e texto.json                   ║
║    python main.py pptx aula.pptx -e texto.json --video           ║
║                                                                  ║
║  vozes                 Lista vozes TTS disponíveis               ║
║    python main.py vozes                                          ║
║                                                                  ║
║  testar-tts            Testa motores TTS                         ║
║    python main.py testar-tts                                     ║
║                                                                  ║
║  OPÇÕES:                                                         ║
║  ───────                                                         ║
║    -c, --config FILE   Ficheiro de configuração                  ║
║    -e, --edicao FILE   Ficheiro JSON/CSV com texto editado       ║
║    --video             Também gerar vídeo MP4                    ║
║    --extrair           Apenas extrair texto para edição          ║
║    -h, --help          Mostrar esta ajuda                        ║
║                                                                  ║
║  FLUXO PPTX RECOMENDADO:                                         ║
║  ───────────────────────                                         ║
║    1. python main.py pptx aula.pptx --extrair                    ║
║    2. Editar ficheiro JSON (ajustar narração)                    ║
║    3. python main.py pptx aula.pptx -e aula_texto.json           ║
║    4. python main.py pptx aula.pptx -e texto.json --video        ║
║                                                                  ║
║  PASTAS:                                                         ║
║  ───────                                                         ║
║    FicheirosEntradaAudio/  - Áudio a converter                   ║
║    FicheirosConvertidos/   - MP3/MP4 gerados                     ║
║    Apresentacoes/          - Ficheiros PPTX                      ║
║    AudioTTS/               - Áudios TTS                          ║
║    Intermedios/            - Ficheiros de edição                 ║
║    img/                    - Imagens para vídeos                 ║
║                                                                  ║
╚══════════════════════════════════════════════════════════════════╝
    """
    print(ajuda)


def main():
    """Função principal - parser de argumentos e execução."""
    parser = argparse.ArgumentParser(
        description="Conversor Multimédia Educativo",
        add_help=False
    )
    
    parser.add_argument(
        "comando",
        nargs="?",
        choices=["audio", "pptx", "vozes", "testar-tts", "help"],
        default="help",
        help="Comando a executar"
    )
    
    parser.add_argument(
        "ficheiro",
        nargs="?",
        help="Ficheiro PPTX a processar"
    )
    
    parser.add_argument(
        "-c", "--config",
        default="config.ini",
        help="Ficheiro de configuração"
    )
    
    parser.add_argument(
        "-e", "--edicao",
        help="Ficheiro JSON/CSV com texto editado"
    )
    
    parser.add_argument(
        "--video",
        action="store_true",
        help="Também gerar vídeo MP4"
    )
    
    parser.add_argument(
        "--extrair",
        action="store_true",
        help="Apenas extrair texto para edição"
    )
    
    parser.add_argument(
        "-h", "--help",
        action="store_true",
        help="Mostrar ajuda"
    )
    
    args = parser.parse_args()
    
    # Executar comando
    if args.help or args.comando == "help":
        mostrar_ajuda()
    
    elif args.comando == "audio":
        processar_audio(args.config)
    
    elif args.comando == "pptx":
        if not args.ficheiro:
            print("❌ Especifique o ficheiro PPTX")
            print("   Exemplo: python main.py pptx apresentacao.pptx")
            sys.exit(1)
        processar_pptx_comando(
            args.ficheiro,
            args.config,
            args.edicao,
            args.video,
            args.extrair
        )
    
    elif args.comando == "vozes":
        listar_vozes_disponiveis()
    
    elif args.comando == "testar-tts":
        testar_tts()
    
    else:
        mostrar_ajuda()


if __name__ == "__main__":
    main()
