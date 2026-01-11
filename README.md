# CHECK LIST DE EQUIPAMENTOS - Centro Cirúrgico

Este repositório contém um script Python para gerar uma planilha Excel de checklist de equipamentos para o Centro Cirúrgico - Sala 01.

## Estrutura da Planilha

A planilha gerada possui:

- **Título**: "CHECK LIST DE EQUIPAMENTOS – CENTRO CIRÚRGICO – SALA 01" (centralizado)
- **Campo de data**: "MÊS/ANO: __________"
- **Tabela com**:
  - Coluna MATERIAL com 15 equipamentos
  - 31 colunas (dias do mês)
  - Cada dia dividido em 3 turnos: M (manhã), T (tarde), N (noite)
  - Linha final para "VISTO DO PROFISSIONAL"
- **Orientação**: Paisagem (horizontal)

## Equipamentos Incluídos

1. ELETROCAUTÉRIO
2. MESA CIRÚRGICA TOMADA
3. CARRO ANEST./CAPNÓGRAFO
4. MONITOR COMPLETO
5. OX PERFURO CORTANTE
6. OX E LÂMINAS LARINGO
7. ESTOJO AD E INF
8. MÓDULO P.A.I
9. PAREDE GASES FUNCIONANDO
10. CIRCUITO ANESTESIA /KTS
11. 02 SUPORTE SORO
12. 03 MESAS + 01 MAYO
13. REANIMADOR AD E INF.
14. 02 LIXEIRAS
15. HAMPER

## Como Usar

### Pré-requisitos

- Python 3.x
- Biblioteca openpyxl

### Instalação

```bash
pip install openpyxl
```

### Gerar a Planilha

```bash
python3 gerar_checklist.py
```

Este comando criará o arquivo `CHECK_LIST_EQUIPAMENTOS_SALA_01.xlsx` no diretório atual.

### Verificar a Planilha

Para verificar se a planilha foi gerada corretamente:

```bash
python3 verificar_checklist.py
```

## Arquivo Gerado

- **Nome**: `CHECK_LIST_EQUIPAMENTOS_SALA_01.xlsx`
- **Formato**: Excel (XLSX)
- **Orientação**: Paisagem
- **Tamanho aproximado**: 11KB

## Características da Planilha

- Todas as células possuem bordas finas
- Cabeçalhos em negrito
- Células de checagem (M, T, N) com largura reduzida para economizar espaço
- Coluna MATERIAL mais larga para acomodar os nomes dos equipamentos
- Formatação centralizada para melhor visualização
- Pronta para impressão em formato A4 paisagem