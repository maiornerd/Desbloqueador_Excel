# 🔓 Desbloqueador de Planilhas Excel

**Programa em Python com interface gráfica para desbloquear planilhas Excel protegidas**

## 📦 Arquivos Inclusos

| Arquivo | Descrição |
|---------|-----------|
| `unlock_excel.py` | 🐍 Programa principal em Python |
| `setup_build.py` | 🔨 Script para criar o executável |
| `instalar.bat` | ⚡ Instalador automático (Windows) |
| `requirements.txt` | 📋 Dependências do projeto |
| `GUIA_INSTALACAO.md` | 📖 Guia completo de instalação |

---

## ⚡ Início Rápido (Windows)

### Opção 1: Instalação Automática (RECOMENDADO)

1. **Duplo clique em `instalar.bat`**
2. Aguarde a conclusão
3. O arquivo `.exe` estará em `dist/Desbloqueador Excel.exe`

### Opção 2: Instalação Manual

```bash
# Passo 1: Instalar dependências
pip install -r requirements.txt

# Passo 2: Criar executável
python setup_build.py

# Passo 3: Executar
dist\Desbloqueador Excel.exe
```

---

## 🎯 Como Usar o Programa

1. **Abra o executável**: `Desbloqueador Excel.exe`
2. **Carregue um arquivo Excel** clicando em "📁 Carregar Arquivo"
3. **Clique em "🔓 Desbloquear Planilha"** para quebrar a proteção
4. **Salve o arquivo** clicando em "💾 Salvar Arquivo"
5. ✅ Pronto! Seu arquivo está desbloqueado

---

## 📋 Requisitos

- **Windows 7 ou superior**
- **Nenhuma instalação adicional necessária** (o .exe é standalone)

---

## 🔧 Para Desenvolvedores

### Modificar o Programa

1. Edite `unlock_excel.py`
2. Execute para testar: `python unlock_excel.py`
3. Recrie o executável quando tiver pronto

### Adicionar Ícone Customizado

```bash
# Coloque um arquivo "icon.ico" na mesma pasta
pyinstaller --onefile --windowed --icon=icon.ico --name="Desbloqueador Excel" unlock_excel.py
```

### Compilar em Outro Sistema

```bash
# Linux/Mac
python3 setup_build.py

# Ou com PyInstaller direto
pyinstaller --onefile --windowed --name="Desbloqueador Excel" unlock_excel.py
```

---

## 📊 Especificações

| Item | Detalhes |
|------|----------|
| **Linguagem** | Python 3.7+ |
| **Interface** | Tkinter (nativa do Python) |
| **Bibliotecas** | openpyxl, pillow |
| **Tamanho .exe** | ~70-80 MB |
| **Compatibilidade** | Windows 7, 8, 10, 11 |
| **Permissões** | Usuário comum (sem admin) |

---

## 🚀 Distribuição para Terceiros

### Copiar Apenas o Executável
```bash
# O arquivo .exe funciona independentemente em qualquer máquina
copy "dist/Desbloqueador Excel.exe" "C:/Seu/Caminho"
```

### Comprimir para Email
```bash
# Comprima apenas o arquivo .exe
# Tamanho final: ~25-30 MB (compactado)
```

### Compartilhar em Rede
1. Coloque o `.exe` em uma pasta compartilhada
2. Crie um atalho para cada usuário
3. Ou distribua via e-mail

---

## ❓ FAQ

**P: Preciso instalar Python em cada máquina?**
R: Não! O executável já contém tudo necessário.

**P: Por que é tão grande (~70MB)?**
R: Porque contém todo o Python + bibliotecas + dependências.

**P: Funciona em Mac/Linux?**
R: Sim, mas precisa recompilar. Use `python3 setup_build.py`

**P: Pode desbloquear qualquer Excel?**
R: Sim, contanto que tenha proteção de planilha (não senha de arquivo).

**P: O programa é seguro?**
R: Sim, é código aberto e executável localmente. Sem conexão externa.

---

## 🐛 Solução de Problemas

### ❌ "instalar.bat" não funciona
- Clique com botão direito → "Executar como administrador"
- Ou copie o caminho da pasta na barra de endereços do explorador

### ❌ "Python não foi encontrado"
- Reinstale Python de https://www.python.org/downloads/
- Marque "Add Python to PATH"

### ❌ "ModuleNotFoundError"
```bash
pip install --upgrade openpyxl pillow pyinstaller
```

### ❌ ".exe demora muito ou trava
- Na primeira execução é normal (lentidão de 3-5 segundos)
- Aguarde o carregamento das dependências

---

## 📝 Notas Importantes

⚠️ **Segurança:**
- Este programa contém heurística de força bruta
- Use apenas em seus próprios arquivos
- Respeite as leis de privacidade e segurança

⚠️ **Antivírus:**
- Alguns antivírus podem flagar como falso positivo
- Configure exceções no seu antivírus se necessário
- O código é 100% seguro e sem malware

---

## 📞 Suporte

Para dúvidas ou problemas:
1. Consulte `GUIA_INSTALACAO.md` para instruções detalhadas
2. Verifique se Python 3.7+ está instalado
3. Tente recriar o executável

---

## 📄 Licença

Este projeto é fornecido como está, sem garantias.

---

## ✨ Versão

**v1.0** - Novembro 2024

---

**Pronto para usar! 🚀**
