# -*- coding: utf-8 -*-
"""
Script de configuração para criar o executável
Execute: pyinstaller --onefile --windowed --icon=icon.ico --name="Desbloqueador Excel" unlock_excel.py
Ou use este script: python setup_build.py
"""

import PyInstaller.__main__
import os
import sys

def criar_executavel():
    """Cria o executável usando PyInstaller"""
    
    print("=" * 60)
    print("CRIANDO EXECUTÁVEL - Desbloqueador de Planilhas Excel")
    print("=" * 60)
    
    # Argumentos para PyInstaller
    args = [
        'unlock_excel.py',
        '--onefile',                    # Cria um único arquivo executável
        '--windowed',                   # Remove janela de console
        '--name=Desbloqueador Excel',   # Nome do executável
        '--specpath=build',             # Pasta para arquivos spec
        '--distpath=dist',              # Pasta para o executável final
        '--workpath=build',             # Pasta de trabalho
    ]
    
    print("\n📦 Iniciando compilação...")
    print(f"Argumentos: {' '.join(args)}\n")
    
    try:
        PyInstaller.__main__.run(args)
        print("\n" + "=" * 60)
        print("✓ SUCESSO! Executável criado com sucesso!")
        print("=" * 60)
        print("\n📁 Arquivo executável localizado em:")
        print("   dist/Desbloqueador Excel.exe")
        print("\n💡 Você pode:")
        print("   1. Executar diretamente: dist/Desbloqueador Excel.exe")
        print("   2. Criar um atalho no Desktop")
        print("   3. Enviar para outras máquinas Windows")
        print("\n⚠️  Nota: Certifique-se que openpyxl está instalado")
        
    except Exception as e:
        print(f"\n✗ ERRO ao criar executável: {str(e)}")
        print("\nTente executar manualmente:")
        print("  pyinstaller --onefile --windowed --name='Desbloqueador Excel' unlock_excel.py")
        sys.exit(1)

if __name__ == "__main__":
    criar_executavel()
