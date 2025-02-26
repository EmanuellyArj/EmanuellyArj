# 👩🏻‍💻 Emanuelly Araújo

**`Analista de Dados`**

Sou uma profissional apaixonada por transformar dados em insights estratégicos, com uma trajetória que une formações e experiências complementares. Minha base acadêmica em Ciências Econômicas e Administração, obtida na UFPR, aliada a um MBA em Finanças e à especialização em Data Science, me proporciona uma visão única e integrada dos negócios. Essa combinação não convencional me permite entender, de forma holística, tanto os desafios operacionais quanto as oportunidades estratégicas que os dados oferecem. Ao longo da minha carreira, tenho atuado na transformação digital de processos, integrando sistemas e criando soluções inovadoras que conectam áreas diversas e promovem uma cultura orientada por dados.

---

### 🤖 Linguagens e Tecnologias

<img 
    align="left" 
    alt="mySQL"
    title="mySQL" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/mysql/mysql-original-wordmark.svg" 
  />
          

<img 
    align="left" 
    alt="SQL"
    title="SQL" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/azuresqldatabase/azuresqldatabase-original.svg" />

<img 
    align="left" 
    alt="Postman"
    title="Postman" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/postman/postman-original.svg"        
  />

<img 
    align="left" 
    alt="dbeaver"
    title="dbeaver" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/dbeaver/dbeaver-original.svg" 
  />
          
  
<img 
    align="left" 
    alt="AWS"
    title="AWS" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/amazonwebservices/amazonwebservices-original-wordmark.svg" 
  />
          

<img 
    align="left" 
    alt="Python" 
    title="Python"
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/python/python-original.svg" 
/>

<img 
    align="left" 
    alt="anaconda"
    title="anaconda" 
    width="30px" 
    style="padding-right: 10px;" 
    src="https://cdn.jsdelivr.net/gh/devicons/devicon@latest/icons/anaconda/anaconda-original-wordmark.svg" 
  />
          

<br/>
<br/>



# Projetos Desenvolvidos


## Automação Contábil para Importações

Este projeto consiste em macros em VBA que automatizam tarefas contábeis no Excel, otimizando o processo de importação e conciliação de dados contábeis. Através da centralização das informações e da automação de processos, o projeto elimina a fragmentação de dados e reduz significativamente o tempo gasto com tarefas manuais e retrabalhos.

## Problema

A contabilização de importações enfrentava diversos desafios, tais como:

- **Fragmentação de Dados:** Informações dispersas em 71 abas, distribuídas em 4 planilhas distintas, dificultando a localização de dados específicos.
- **Processos Manuais Demorados:**  
  - Atualização da taxa de câmbio demandava, em média, 4 horas mensais.  
  - A busca por aquisições antigas consumia cerca de 20 minutos por operação.
- **Retrabalho Acumulado:** De janeiro a agosto, foram registradas 14 horas de retrabalho.
- **Ausência de Regras Documentadas:** Os lançamentos contábeis eram realizados manualmente sem regras formalizadas, o que aumentava o risco de erros.

## Solução

O projeto implementa melhorias significativas por meio da automação com VBA:

- **Centralização de Informações:** Consolidação dos dados em uma única planilha, permitindo a localização imediata dos registros por meio de filtros.
- **Automação de Processos:**  
  - **Macro 1:** Automatiza a criação de lançamentos atualizados, realizando a identificação e processamento de registros, além de atualizar automaticamente as taxas de câmbio.  
  - **Macro 2:** Realiza a conciliação dos lançamentos e gera relatórios prontos para importação no sistema contábil, integrando cálculos, histórico e tratamento de exceções.
- **Eliminação do Retrabalho:** Com a centralização e a automação, todas as atualizações e validações são realizadas de forma rápida e precisa, eliminando a necessidade de processos manuais.

## Macros

### Macro 1: Criação de lançamentos atualizados automáticos

```vba
Sub CriarLinhas()
    Dim wsMov As Worksheet
    Dim wsPTAX As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim ultimaData As Date
    Dim novaLinha As Long
    Dim identificador As String
    Dim tipoLancamento As String
    Dim existeLiquidado As Boolean
    Dim lançamentosAdicionados As Long
    Dim identificadoresProcessados As Object
    Dim dataExistente As Boolean
    Dim taxaPTAX As Double
    Dim dataBusca As Date
    Dim ptaxUltimaLinha As Long
    Dim encontrado As Boolean
    Dim j As Long
    Dim mesAnoFiltro As String
    Dim mesFiltro As Integer
    Dim anoFiltro As Integer
    Dim mesAnoIdentificador As String

    ' Inicializa o contador de lançamentos adicionados
    lançamentosAdicionados = 0

    ' Definindo as referências para as planilhas
    Set wsMov = ThisWorkbook.Sheets("Movimentações")
    Set wsPTAX = ThisWorkbook.Sheets("PTAX")
    
    ' Criando um dicionário para armazenar identificadores processados
    Set identificadoresProcessados = CreateObject("Scripting.Dictionary")

    ' Obter mês/ano da célula C3
    mesAnoFiltro = wsMov.Range("C3").Value
    If mesAnoFiltro = "" Then
        MsgBox "Por favor, insira um mês/ano válido na célula C3 (MM/YYYY).", vbExclamation
        Exit Sub
    End If
    
    ' Validar formato da data
    If Not IsDate(mesAnoFiltro) Then
        MsgBox "Formato inválido em C3. Use MM/YYYY.", vbExclamation
        Exit Sub
    End If
    
    ' Extrair mês e ano
    mesFiltro = Month(CDate(mesAnoFiltro))
    anoFiltro = Year(CDate(mesAnoFiltro))
    
    ' Calcular último dia do mês informado
    ultimaData = DateSerial(anoFiltro, mesFiltro + 1, 0)

    ' Última linha da planilha Movimentações
    ultimaLinha = wsMov.Cells(wsMov.Rows.Count, 1).End(xlUp).Row

    ' Última linha da planilha PTAX
    ptaxUltimaLinha = wsPTAX.Cells(wsPTAX.Rows.Count, 1).End(xlUp).Row

    ' Loop para verificar os identificadores na coluna A
    For i = 7 To ultimaLinha ' Começar da linha 7
        identificador = wsMov.Cells(i, 1).Value
        
        If Not identificadoresProcessados.exists(identificador) Then
            existeLiquidado = False
            dataExistente = False
            
            ' Verificar ocorrências do identificador
            For j = 7 To ultimaLinha
                If wsMov.Cells(j, 1).Value = identificador Then
                    tipoLancamento = wsMov.Cells(j, 13).Value ' Coluna M
                    
                    ' Verificar liquidação
                    If tipoLancamento = "LIQUIDADO" Then
                        existeLiquidado = True
                        wsMov.Cells(j, 17).Value = "Liquidado"
                    End If
                    
                    ' Verificar se já existe linha para o mês/ano informado
                    mesAnoIdentificador = Format(wsMov.Cells(j, 2).Value, "mm/yyyy")
                    If mesAnoIdentificador = Format(ultimaData, "mm/yyyy") Then
                        dataExistente = True
                    End If
                End If
            Next j
            
            If existeLiquidado Then
                ' Marcar todas as linhas do identificador como Liquidado
                For j = 7 To ultimaLinha
                    If wsMov.Cells(j, 1).Value = identificador Then
                        wsMov.Cells(j, 17).Value = "Liquidado"
                    End If
                Next j
            ElseIf Not dataExistente Then
                ' Criar nova linha
                novaLinha = ultimaLinha + 1
                
                ' Copiar identificador e data
                wsMov.Cells(novaLinha, 1).Value = identificador
                wsMov.Cells(novaLinha, 2).Value = ultimaData
                
                ' Buscar PTAX para a data
                dataBusca = ultimaData
                encontrado = False
                For j = 2 To ptaxUltimaLinha
                    If wsPTAX.Cells(j, 1).Value = dataBusca Then
                        taxaPTAX = wsPTAX.Cells(j, 2).Value
                        encontrado = True
                        Exit For
                    End If
                Next j
                
                ' Preencher taxa PTAX ou mensagem de erro
                wsMov.Cells(novaLinha, 10).Value = IIf(encontrado, taxaPTAX, "Data não encontrada")
                
                ' Copiar demais dados
                wsMov.Range("D" & i & ":I" & i).Copy wsMov.Range("D" & novaLinha)
                wsMov.Cells(novaLinha, 11).Value = wsMov.Cells(i, 11).Value
                wsMov.Cells(novaLinha, 13).Value = wsMov.Cells(i, 13).Value
                wsMov.Cells(novaLinha, 14).Value = wsMov.Cells(i, 14).Value
                
                ultimaLinha = ultimaLinha + 1
                lançamentosAdicionados = lançamentosAdicionados + 1
            End If

            identificadoresProcessados.Add identificador, True
        End If
    Next i

    ' Mensagem final
    If lançamentosAdicionados > 0 Then
        MsgBox lançamentosAdicionados & " lançamento(s) adicionado(s) para " & Format(ultimaData, "MM/yyyy") & "!", vbInformation
    Else
        MsgBox "Nenhum lançamento novo necessário para " & Format(ultimaData, "MM/yyyy") & ".", vbInformation
    End If
End Sub

````

## Macro 2: Conciliação e criação de relatório automático

```vba
Sub FiltrarLancamentosAtualizado()
    Dim wsMov As Worksheet
    Dim wsPTAX As Worksheet
    Dim wsContabil As Worksheet
    Dim wsLanc As Worksheet
    Dim mesAnoFiltro As String
    Dim ultimaLinhaMov As Long
    Dim linhaLanc As Long
    Dim linhaMov As Long
    Dim dataMov As Date
    Dim mesFiltro As Integer
    Dim anoFiltro As Integer
    Dim valorCalculado As Double
    Dim identificador As String
    Dim tipoMov As String
    Dim passivoAtivo As String
    Dim ultimaDataPTAX As Date
    Dim valorDolarAnterior As Double
    Dim valorDolarAtual As Double
    Dim estrutura As String
    Dim tipoLiquidado As String
    Dim historico As String
    Dim somaValorUSD As Double
    Dim taxaCambio As Double
    Dim valorReais As Double
    Dim varDolar As String
    Dim valorContaContabil As String
    Dim contaCadastrada As Boolean
    Dim ultimaLinhaContabil As Long
    Dim lookupRange As Range
    Dim resultadoContabil As Variant

    ' Definir as abas
    Set wsMov = ThisWorkbook.Sheets("Movimentações")
    Set wsLanc = ThisWorkbook.Sheets("Lançamento")
    Set wsPTAX = ThisWorkbook.Sheets("PTAX")
    Set wsContabil = ThisWorkbook.Sheets("Conta Contabil")
    
    ' Obter o mês e ano do filtro na célula H3
    mesAnoFiltro = wsLanc.Range("H3").Value
    If mesAnoFiltro = "" Then
        MsgBox "Por favor, insira um mês e ano válido na célula H3.", vbExclamation
        Exit Sub
    End If
    
    ' Validar formato da data
    If Not IsDate(mesAnoFiltro) Then
        MsgBox "Data no formato incorreto. Use MM/YYYY.", vbExclamation
        Exit Sub
    End If
    
    ' Extrair mês e ano
    mesFiltro = Month(CDate(mesAnoFiltro))
    anoFiltro = Year(CDate(mesAnoFiltro))
    
    ' Identificar última linha da aba Movimentações
    ultimaLinhaMov = wsMov.Cells(wsMov.Rows.Count, "A").End(xlUp).Row
    
    ' Identificar intervalo dinâmico na Conta Contabil
    ultimaLinhaContabil = wsContabil.Cells(wsContabil.Rows.Count, "A").End(xlUp).Row
    Set lookupRange = wsContabil.Range("A2:B" & ultimaLinhaContabil) ' Assume que a linha 1 é cabeçalho
    
    ' Limpar dados existentes
    wsLanc.Rows("7:" & wsLanc.Rows.Count).ClearContents
    
    ' Iniciar preenchimento
    linhaLanc = 7

    For linhaMov = 7 To ultimaLinhaMov
        If IsDate(wsMov.Cells(linhaMov, "B").Value) Then
            dataMov = wsMov.Cells(linhaMov, "B").Value
            
            If Month(dataMov) = mesFiltro And Year(dataMov) = anoFiltro Then
                ' Verificar se está liquidado em M ou Q
                If UCase(Trim(wsMov.Cells(linhaMov, "M").Value)) = "LIQUIDADO" Or _
                   UCase(Trim(wsMov.Cells(linhaMov, "Q").Value)) = "LIQUIDADO" Then
                    GoTo ContinueLoop ' Pula linhas liquidadas
                End If
                
                ' Cálculo do valor
                valorCalculado = (wsMov.Cells(linhaMov, "K").Value + wsMov.Cells(linhaMov, "L").Value) * wsMov.Cells(linhaMov, "J").Value
                
                ' IDENTIFICADOR CORRETO (COLUNA A DA MOVIMENTAÇÕES)
                identificador = Trim(CStr(wsMov.Cells(linhaMov, "A").Value))
                tipoMov = wsMov.Cells(linhaMov, "M").Value
                
                ' Determinar Ativo/Passivo (COLUNA N)
                passivoAtivo = UCase(Trim(wsMov.Cells(linhaMov, "N").Value))
                If passivoAtivo <> "ATIVO" And passivoAtivo <> "PASSIVO" Then
                    Debug.Print "Valor inválido na linha " & linhaMov & ": " & passivoAtivo
                End If
                
                ' Depuração reforçada
                Debug.Print "Processando linha " & linhaMov & _
                    " | ID: " & identificador & _
                    " | Tipo: " & passivoAtivo & _
                    " | Mov: " & tipoMov
                
                ' Cálculo PTAX
                ultimaDataPTAX = WorksheetFunction.EoMonth(CDate(dataMov), -1)
                
                ' Busca valores dólar
                On Error Resume Next
                If passivoAtivo = "ATIVO" Then
                    valorDolarAnterior = Application.VLookup(ultimaDataPTAX, wsPTAX.Range("A:B"), 2, False)
                Else
                    valorDolarAnterior = Application.VLookup(ultimaDataPTAX, wsPTAX.Range("A:C"), 3, False)
                End If
                On Error GoTo 0
                
                ' Tratamento de erros PTAX
                If IsError(valorDolarAnterior) Then
                    valorDolarAnterior = 0
                    Debug.Print "Erro PTAX na linha " & linhaMov
                End If
                
                valorDolarAtual = wsMov.Cells(linhaMov, "J").Value
                If valorDolarAtual = 0 Then Debug.Print "PTAX zero na linha " & linhaMov
                
                ' Determinar variação
                varDolar = IIf(valorDolarAtual > valorDolarAnterior, "aumentou", _
                             IIf(valorDolarAtual < valorDolarAnterior, "diminuiu", "estável"))
                
                ' Definir estrutura
                estrutura = Switch(passivoAtivo = "ATIVO", "132", passivoAtivo = "PASSIVO", "133")
                
                ' Preencher dados básicos
                With wsLanc
                    .Cells(linhaLanc, "A").Value = identificador
                    .Cells(linhaLanc, "B").Value = dataMov
                    .Cells(linhaLanc, "E").Value = valorCalculado
                    .Cells(linhaLanc, "F").Value = varDolar
                    .Cells(linhaLanc, "G").Value = estrutura
                    .Cells(linhaLanc, "H").Value = IIf(estrutura = "133", _
                        "VLR VARIAÇÃO CAMBIAL PASSIVA - PROVISÃO", _
                        "VLR VARIAÇÃO CAMBIAL ATIVA - PROVISÃO")
                End With
                
                ' Montar histórico
                somaValorUSD = wsMov.Cells(linhaMov, "K").Value + wsMov.Cells(linhaMov, "L").Value
                taxaCambio = valorDolarAtual
                valorReais = somaValorUSD * taxaCambio
                
                historico = ""
                With wsMov
                    If .Cells(linhaMov, "I").Value <> "" Then historico = historico & "FORNECEDOR " & .Cells(linhaMov, "I").Value & " "
                    If .Cells(linhaMov, "D").Value <> "" Then historico = historico & "CONTRATO " & .Cells(linhaMov, "D").Value & " "
                    If .Cells(linhaMov, "E").Value <> "" Then historico = historico & "INVOICE " & .Cells(linhaMov, "E").Value & " "
                    If .Cells(linhaMov, "F").Value <> "" Then historico = historico & "DI " & .Cells(linhaMov, "F").Value & " "
                    If .Cells(linhaMov, "G").Value <> "" Then historico = historico & "NF " & .Cells(linhaMov, "G").Value & " "
                    If .Cells(linhaMov, "H").Value <> "" Then historico = historico & "FINIMP " & .Cells(linhaMov, "H").Value & " "
                End With
                
                historico = historico & "REF USD " & Format(somaValorUSD, "#,##0.00") & _
                            " TAXA USD " & Format(taxaCambio, "#,##0.0000") & _
                            " = R$ " & Format(valorReais, "#,##0.00")
                
                wsLanc.Cells(linhaLanc, "I").Value = historico
                
                ' PREENCHIMENTO DAS COLUNAS C E D (BUSCA PELO ID CORRETO)
                contaCadastrada = False
                valorContaContabil = ""
                
                ' Nova lógica de busca com tratamento robusto
                On Error Resume Next ' Ignorar erros temporariamente
                
                ' Verificar se o identificador é numérico
                If IsNumeric(identificador) Then
                    ' Buscar como número
                    resultadoContabil = Application.VLookup(CLng(identificador), lookupRange, 2, False)
                Else
                    ' Buscar como texto
                    resultadoContabil = Application.VLookup(identificador, lookupRange, 2, False)
                End If
                
                On Error GoTo 0 ' Restaurar tratamento de erros
                
                If Not IsError(resultadoContabil) Then
                    valorContaContabil = CStr(resultadoContabil)
                    contaCadastrada = True
                    Debug.Print "Linha " & linhaMov & ": Conta encontrada - " & valorContaContabil
                Else
                    Debug.Print "Linha " & linhaMov & ": Conta NÃO encontrada para ID " & identificador & _
                             " (Tipo: " & TypeName(identificador) & ")"
                End If
                
                ' Lógica complexa para Débito/Crédito
                Select Case passivoAtivo
                    Case "ATIVO"
                        If tipoMov = "ADIANTAMENTO" Then
                            If varDolar = "aumentou" Then
                                wsLanc.Cells(linhaLanc, "C").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                                wsLanc.Cells(linhaLanc, "D").Value = 2666
                            Else
                                wsLanc.Cells(linhaLanc, "C").Value = 2356
                                wsLanc.Cells(linhaLanc, "D").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                            End If
                        End If
                        
                    Case "PASSIVO"
                        Select Case tipoMov
                            Case "FORNECEDOR"
                                If varDolar = "aumentou" Then
                                    wsLanc.Cells(linhaLanc, "C").Value = 2356
                                    wsLanc.Cells(linhaLanc, "D").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                                Else
                                    wsLanc.Cells(linhaLanc, "C").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                                    wsLanc.Cells(linhaLanc, "D").Value = 2666
                                End If
                                
                            Case "FINIMP"
                                If varDolar = "aumentou" Then
                                    wsLanc.Cells(linhaLanc, "C").Value = IIf(contaCadastrada, valorContaContabil, "Cadastrar Conta")
                                    wsLanc.Cells(linhaLanc, "D").Value = valorContaContabil
                                Else
                                    wsLanc.Cells(linhaLanc, "C").Value = 2666
                                    wsLanc.Cells(linhaLanc, "D").Value = valorContaContabil
                                End If
                        End Select
                End Select
                
                ' Atualizar Movimentações
                wsMov.Cells(linhaMov, "O").Value = wsLanc.Cells(linhaLanc, "C").Value
                wsMov.Cells(linhaMov, "P").Value = wsLanc.Cells(linhaLanc, "D").Value
                
                linhaLanc = linhaLanc + 1
            End If
        End If
ContinueLoop:
    Next linhaMov
    
    MsgBox "Processamento concluído!", vbInformation

    ' CHAMAR A MACRO PARA GERAR O ARQUIVO EXCEL
    Call CriarArquivoExcel
End Sub

````

## Conciliação Fiscal: Tributação de ICMS e IPI 

Este projeto consiste em macros em VBA que automatizam tarefas fiscais no Excel, otimizando o processo de importação e conciliação de dados contábeis. Através da automação de processos, o projeto elimina a fragmentação de dados e reduz significativamente o tempo gasto com tarefas manuais e retrabalhos.

## Problema

A conciliação de impostos de ICMS e IPI enfrentava diversos desafios, tais como:

- **Processos Manuais Demorados:**  
  - A conciliação demandava, em média, 2 hora mensal por empresa.
- **Retrabalho :** Para evitar erros e multas esse processo era revisado por no 2 pessoas, o que se tornava muito moroso e custoso.
- **Ausência de Regras Documentadas:** Os lançamentos contábeis eram realizados manualmente sem regras formalizadas, o que aumentava o risco de erros.
- **Multas :** Por falta de padronização, acabavam acontecendo erros, o que gerava um volume alto de perdas financeiras

## Solução

O projeto implementa melhorias significativas por meio da automação com VBA:

- **Automação de Processos:**  
  - **Macro 1:** Automatiza a criação de lançamentos atualizados, realizando a identificação e processamento de registros, além de atualizar automaticamente as taxas de câmbio.  
  - **Macro 2:** Realiza a conciliação dos lançamentos e gera relatórios prontos para importação no sistema contábil, integrando cálculos, histórico e tratamento de exceções.
- **Eliminação do Retrabalho:** Com a centralização e a automação, todas as atualizações e validações são realizadas de forma rápida e precisa, eliminando a necessidade de processos manuais.

## Macros

### Macro 1: Criação de lançamentos atualizados automáticos

```vba
