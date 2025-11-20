# 🔓 Guia Completo - Desbloqueador de Planilhas Excel

## 📋 Pré-requisitos

Antes de criar o executável, certifique-se de ter:

1. **Python 3.7 ou superior instalado** ([https://www.python.org/downloads/](https://www.python.org/downloads/))
   - ✅ Marque "Add Python to PATH" durante a instalação

2. **Git (opcional, mas recomendado)** ([https://git-scm.com/download/win](https://git-scm.com/download/win))

---

## 🚀 Método 1: Criar Executável (Recomendado)

### Passo 1: Preparar a Pasta
```bash
# Crie uma pasta para o projeto
mkdir Desbloqueador-Excel
cd Desbloqueador-Excel
```

### Passo 2: Copiar Arquivos
Coloque nesta pasta os arquivos:
- `unlock_excel.py` (o programa principal)
- `setup_build.py` (script para criar o executável)
- `requirements.txt` (lista de dependências)

### Passo 3: Instalar Dependências
```bash
pip install -r requirements.txt
```

Ou manualmente:
```bash
pip install openpyxl pillow pyinstaller
```

### Passo 4: Criar o Executável

**Opção A: Usar o script automatizado**
```bash
python setup_build.py
```

**Opção B: Comando direto do PyInstaller**
```bash
pyinstaller --onefile --windowed --name="Desbloqueador Excel" unlock_excel.py
```

### Passo 5: Encontrar o Executável

Após a compilação, o arquivo estará em:
```
dist/Desbloqueador Excel.exe
```

---

## 📦 Método 2: Criar Instalador (Avançado)

Se desejar criar um instalador profissional (.msi):

### Passo 1: Instalar o NSIS
```bash
pip install pyinstaller nsis
```

### Passo 2: Criar o Instalador
```bash
pyinstaller --onefile --windowed --name="Desbloqueador Excel" \
  --distpath=dist_installer unlock_excel.py

# Copie o arquivo para a pasta de instalação
```

---

## ✅ Método 3: Solução Pré-compilada

Se você tiver dificuldades em compilar, pode:

1. **Solicitar o .exe pré-compilado** ao desenvolvedor
2. **Usar plataformas online** como:
   - [PyPI](https://pypi.org/)
   - [py2exe.org](http://py2exe.org/)

---

## 🎯 Como Usar o Executável

### Na Máquina de Desenvolvimento:
1. Execute: `dist/Desbloqueador Excel.exe`
2. A interface gráfica será aberta

### Em Outra Máquina (Distribuição):
1. Copie apenas o arquivo `.exe` para a outra máquina
2. ✅ Não precisa instalar Python!
3. Execute o `.exe` diretamente

### Criar Atalho no Desktop:
1. Clique com botão direito no `.exe`
2. Selecione "Criar Atalho"
3. Mova o atalho para o Desktop

---

## 🔧 Solução de Problemas

### ❌ "ModuleNotFoundError: No module named 'openpyxl'"
```bash
pip install openpyxl
```

### ❌ "Python não reconhecido"
- Reinstale Python com "Add Python to PATH" marcado
- Ou execute no terminal administrativo

### ❌ ".exe não funciona em outra máquina"
- Certifique-se de ter copiado APENAS o arquivo `.exe`
- Teste com `--windowed` flag no PyInstaller

### ❌ "Arquivo muito grande"
Use a flag `--onefile` para reduzir o tamanho:
```bash
pyinstaller --onefile --windowed unlock_excel.py
```

---

## 📊 Tamanho do Arquivo Final

- **Com --onefile**: ~60-80 MB
- **Sem --onefile**: ~10-15 MB (múltiplos arquivos)

---

## 🔐 Segurança

⚠️ **Importante:**
- O executável contém todas as dependências Python
- Pode ser detectado como falso positivo por antivírus (normal)
- Se isto acontecer, configure exceções no seu antivírus

---

## 📝 Estrutura de Arquivos Final

```
Desbloqueador-Excel/
├── unlock_excel.py           (Programa principal)
├── setup_build.py            (Script de compilação)
├── requirements.txt          (Dependências)
├── build/                    (Arquivos temporários - pode deletar)
├── dist/                     (EXECUTÁVEL FINAL)
│   └── Desbloqueador Excel.exe
└── README.md                 (Este arquivo)
```

---

## 🎁 Distribuição

### Para compartilhar com colegas:

1. **Opção A: Apenas o .exe**
   - Mais simples
   - Arquivo ~70 MB
   - Não precisa Python instalado

2. **Opção B: Pasta completa**
   - Include tudo
   - Fácil para atualizar
   - Permite customizações

3. **Opção C: Criar um instalador**
   - Profissional
   - Automático
   - Gerencia desinstalação

---

## 📞 Suporte

Se encontrar problemas:
1. Verifique se Python 3.7+ está instalado
2. Reinstale as dependências: `pip install --upgrade openpyxl pillow`
3. Tente recriar o executável
4. Consulte a documentação do PyInstaller: [https://pyinstaller.org/](https://pyinstaller.org/)

---

## ✨ Próximos Passos

- ✅ Executável criado e testado
- ✅ Pronto para distribuição
- ✅ Funciona em qualquer PC Windows

Bom uso! 🚀
