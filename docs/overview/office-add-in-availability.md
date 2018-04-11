---
title: Disponibilidade de host e plataforma para suplementos do Office
description: 'Conjuntos de requisitos compatíveis com Excel, Word, Outlook, PowerPoint e OneNote.'
ms.date: 3/20/2018
---

# <a name="office-add-in-host-and-platform-availability"></a>Disponibilidade de host e plataforma para suplementos do Office

Seu suplemento do Office pode depender de um host específico do Office, um conjunto de requisitos, um membro de API ou uma versão da API para funcionar conforme o esperado. As tabelas a seguir contêm as plataformas disponíveis, os pontos de extensão, os conjuntos de requisitos de API e os conjuntos de requisitos comuns de API que atualmente são compatíveis com cada aplicativo do Office. 

Se uma célula de tabela apresenta um asterisco (*), significa que estamos trabalhando no assunto. Confira os conjuntos de requisitos do Project ou do Access em [Conjuntos de requisitos comuns do Office](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets).  

> [!NOTE]
> O número do build do Office 2016 instalado via MSI é 16.0.4266.1001. Esta versão só contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API.

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Plataforma</th>
    <th style="width:10%">Pontos de extensão</th> 
    <th style="width:20%">Conjuntos de requisitos da API</th> 
    <th style="width:40%"><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Painel de tarefas<br>
        - Conteúdo<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a>
    </td>
    <td>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - CompressedFile<br>
        - Configurações<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 para Windows</td>
    <td>
        - Painel de tarefas<br>
        - Conteúdo</td>
    <td>  - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>
        - BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Configurações<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 para Windows</td>
    <td>- Painel de tarefas<br>
        - Conteúdo<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.4</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Configurações<br>
        - TextCoercion</td> 
  </tr>
  <tr>
    <td>Office para iOS</td>
    <td>- Painel de tarefas<br>
        - Conteúdo</td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - Configurações<br>
        - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 para Mac</td>
    <td>- Painel de tarefas<br>
        - Conteúdo<br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td>- <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/excel-api-requirement-sets">ExcelApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td>- BindingEvents<br>
        - DocumentEvents<br>
        - MatrixBindings<br>
        - MatrixCoercion<br>
        - TableBindings<br>
        - TableCoercion<br>
        - TextBindings<br>
        - TextCoercion</td>
  </tr>
</table>

<br/>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th> 
    <th>Conjuntos de requisitos da API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></th> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Leitura de email<br>
      - Composição de email<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1"> Caixa de correio 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3"> Caixa de correio 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4"> Caixa de correio 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5"> Caixa de correio 1.5</a></td>
    <td>não disponível</td>
  </tr>
  <tr>
    <td>Office 2013 para Windows</td>
    <td> - Leitura de email<br>
      - Composição de email<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1"> Caixa de correio 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3"> Caixa de correio 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4"> Caixa de correio 1.4</a></td>
    <td>não disponível</td>
  </tr>
  <tr>
    <td>Office 2016 para Windows</td>
    <td> - Leitura de email<br>
      - Composição de email<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a><br>
      - Módulos</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1"> Caixa de correio 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3"> Caixa de correio 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4"> Caixa de correio 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5"> Caixa de correio 1.5</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.6/index?product=outlook&version=v1.6">Caixa de correio 1.6</a></td>
    <td>não disponível</td> 
  </tr>
  <tr>
    <td>Office para iOS</td>
    <td> - Leitura de email<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1"> Caixa de correio 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3"> Caixa de correio 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4"> Caixa de correio 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5"> Caixa de correio 1.5</a></td>    
    <td>não disponível</td>
  </tr>
  <tr>
    <td>Office 2016 para Mac</td>
    <td> - Leitura de email<br>
      - Composição de email<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1"> Caixa de correio 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3"> Caixa de correio 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4"> Caixa de correio 1.4</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.5/index?product=outlook&version=v1.5"> Caixa de correio 1.5</a></td>
    <td>não disponível</td>
  </tr>
  <tr>
    <td>Office para Android</td>
    <td> - Leitura de email<br>
      - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/outlook/1.1/index?product=outlook&version=v1.1"> Caixa de correio 1.1</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.2/index?product=outlook&version=v1.2">Caixa de correio 1.2</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.3/index?product=outlook&version=v1.3"> Caixa de correio 1.3</a><br>
      - <a href="https://dev.office.com/reference/add-ins/outlook/1.4/index?product=outlook&version=v1.4"> Caixa de correio 1.4</a></td>
    <td>não disponível</td>
  </tr>
</table>

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th> 
    <th>Conjuntos de requisitos da API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Painel de tarefas<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CustomXmlParts<br>
         - MatrixBindings<br>
         - MatrixCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - DocumentEvents<br>
         - TextFile<br>
         - ImageCoercion<br>
         - Configurações<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 para Windows</td>
    <td> - Painel de tarefas</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - Arquivo<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Configurações<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - Associações de matriz</td>
  </tr>
  <tr>
    <td>Office 2016 para Windows</td>
    <td> - Painel de tarefas<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - Arquivo<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Configurações<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - Associações de matriz </td> 
  </tr>
  <tr>
    <td>Office para iOS</td>
    <td> - Painel de tarefas</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - Arquivo<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Configurações<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - Associações de matriz </td> 
  </tr>
  <tr>
    <td>Office 2016 para Mac</td>
    <td> - Painel de tarefas<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.1</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.2</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/word-api-requirement-sets">WordApi 1.3</a><br>
        - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - BindingEvents<br>
         - CompressedFile<br>
         - CustomXmlPart<br>
         - DocumentEvents<br>
         - Arquivo<br>
         - HtmlCoercion<br>
         - ImageCoercion<br>
         - OoxmlCoercion<br>
         - TableBindings<br>
         - TableCoercion<br>
         - TextBindings<br>
         - TextFile<br>
         - Configurações<br>
         - TextCoercion<br>
         - MatrixCoercion<br>
         - Associações de matriz </td> 
  </tr>
</table>

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th> 
    <th>Conjuntos de requisitos da API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Conteúdo<br>
         - Painel de tarefas<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - Arquivo<br>
         - Seleção<br>
         - Configurações<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 para Windows</td>
    <td> - Conteúdo<br>
         - Painel de tarefas<br>
    </td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a>
</td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - Arquivo<br>
         - Seleção<br>
         - Configurações<br>
         - TextCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 para Windows</td>
    <td> - Conteúdo<br>
         - Painel de tarefas<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - Arquivo<br>
         - Seleção<br>
         - Configurações<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office para iOS</td>
    <td> - Conteúdo<br>
         - Painel de tarefas</td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
     <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - Arquivo<br>
         - Seleção<br>
         - Configurações<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2016 para Mac</td>
    <td> - Conteúdo<br>
         - Painel de tarefas<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - ActiveView<br>
         - CompressedFile<br>
         - DocumentEvents<br>
         - Arquivo<br>
         - Seleção<br>
         - Configurações<br>
         - TextCoercion<br>
         - ImageCoercion</td>
  </tr>
</table>

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th> 
    <th>Conjuntos de requisitos da API</th> 
    <th><a href="https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets"><b>APIs comuns</b></a></th> 
  </tr> 
  </tr>
  <tr>
    <td>Office Online</td>
    <td> - Conteúdo<br>
         - Painel de tarefas<br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets">Comandos de suplemento</a></td>
    <td> - <a href="https://dev.office.com/reference/add-ins/requirement-sets/onenote-api-requirement-sets">OneNoteApi 1.1</a><br>
         - <a href="https://dev.office.com/reference/add-ins/requirement-sets/dialog-api-requirement-sets">DialogApi 1.1</a></td>
    <td> - DocumentEvents<br>
         - Configurações<br>
         - TextCoercion<br>
         - HtmlCoercion<br>
         - ImageCoercion</td>
  </tr>
  <tr>
    <td>Office 2013 para Windows</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr> 
  <tr>
    <td>Office 2016 para Windows</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td> 
  </tr>
  <tr>
    <td>Office para iOS</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
  <tr>
    <td>Office 2016 para Mac</td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
    <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* </td>
  </tr>
</table>

<br/>

\* = Estamos trabalhando nisso. 

## <a name="see-also"></a>Veja também

- [Visão geral da plataforma Suplementos do Office](office-add-ins.md)
- [Conjuntos de requisitos comuns da API](https://dev.office.com/reference/add-ins/requirement-sets/office-add-in-requirement-sets)
- [Conjuntos de requisitos dos comandos de suplemento](https://dev.office.com/reference/add-ins/requirement-sets/add-in-commands-requirement-sets)
- [Referência da API JavaScript do Office](https://dev.office.com/reference/add-ins/javascript-api-for-office)

