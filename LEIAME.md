# Planilha Financeira Avançada (pessoa física)

Arquivos entregues:

| Arquivo | O que é |
|---|---|
| `Planilha_Financeira_Avancada.xlsx` | Planilha completa **com dados de EXEMPLO** fictícios (marcados "EXEMPLO") |
| `Planilha_Financeira_Avancada_LIMPA.xlsx` | A mesma planilha, **sem exemplos**, pronta para os seus dados |
| `gerador/` | Código Python (openpyxl) que gera as duas planilhas (`gerador/finalizar.sh` ou `gerador/build.py`) |

Comece pela aba **COMEÇAR AQUI**. No dia a dia você só digita em **LANÇAMENTOS**; o resto é calculado.
Células azuis = entrada · cinza = calculada (protegida, sem senha).

Veja a aba **BACKUP** (premissas e limitações) e o final da aba **COMEÇAR AQUI** (como limpar os exemplos).

## Regerar

```bash
cd gerador
./finalizar.sh        # gera as duas planilhas, recalcula no LibreOffice (headless), injeta o cache e confere erros
```
