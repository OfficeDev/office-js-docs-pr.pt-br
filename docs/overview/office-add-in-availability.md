---
title: Disponibilidade de aplicativos e plataformas do cliente Office para Suplementos do Office
description: Conjuntos de requisitos com suporte para o Excel, OneNote, Outlook, PowerPoint, Project e Word.
ms.date: 07/13/2021
localization_priority: Priority
ms.openlocfilehash: 7b3bd770d74f29d1a0b27da5080284aa62146101
ms.sourcegitcommit: 30a861ece18255e342725e31c47f01960b854532
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/16/2021
ms.locfileid: "53455492"
---
# <a name="office-client-application-and-platform-availability-for-office-add-ins"></a>Disponibilidade de aplicativos e plataformas do cliente Office para Suplementos do Office

Para funcionar conforme o esperado, o Suplemento do Office pode depender de um aplicativo específico do Office, um conjunto de requisitos, um membro da API ou uma versão da API. As tabelas a seguir contêm as plataformas disponíveis, pontos de extensão, conjuntos de requisitos de API e APIs comuns que são atualmente suportados para cada aplicativo do Office.

<br>

|<a href="#excel"><img src="../images/index/logo-excel.svg" alt="Excel" width="48" /><br><span>Excel</span></a>|<a href="#onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote" width="48" /><br><span>OneNote</span></a>|<a href="#outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook" width="48" /><br><span>Outlook</span></a>|<a href="#powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint" width="48" /><br><span>PowerPoint</span></a>|<a href="#project"><img src="../images/index/logo-project-server.svg" alt="Project" width="48" /><br><span>Project</span></a>|<a href="#word"><img src="../images/index/logo-word.svg" alt="Word" width="48" /><br><span>Word</span></a>|
|:---:|:---:|:---:|:---:|:---:|:---:|

> [!NOTE]
> A versão inicial do Office 2016 instalada por meio do MSI apenas contém os conjuntos de requisitos ExcelApi 1.1, WordApi 1.1 e os conjuntos de requisitos comuns de API. Para saber mais sobre o histórico de atualizações de várias versões do Office, confira a seção[Confira também](#see-also). Os Suplementos do Office podem não ter suporte em todos os serviços que são membros do [Programa de Parceiros de Armazenamento em Nuvem do Office](https://developer.microsoft.com/office/cloud-storage-partner-program), que permite a integração do Office na Web para trabalhar com documentos do Office como parte de sua oferta de serviço. Para obter mais informações, entre em contato com o serviço de membro.

## <a name="excel"></a>Excel

<table style="width:80%">
  <tr>
    <th style="width:10%">Plataforma</th>
    <th style="width:10%">Pontos de extensão</th>
    <th style="width:20%">Conjuntos de requisitos da API</th>
    <th style="width:40%"><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></th>
  </tr>
  <tr>
    <td>Office na Web</td>
    <td>
      - TaskPane<br>
      - Conteúdo<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/excel-api-online-requirement-set.md">ExcelApiOnline</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office no Windows<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - TaskPane<br>
      - Conteúdo<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 no Windows<br>(compra avulsa)</td>
    <td>
      - TaskPane<br>
      - Conteúdo<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 no Windows<br>(compra avulsa)</td>
    <td>
      - TaskPane<br>
      - Conteúdo </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 no Windows<br>(compra avulsa)</td>
    <td>
      - TaskPane<br>
      - Conteúdo </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office no iPad<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - TaskPane<br>
      - Conteúdo </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office no Mac<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - TaskPane<br>
      - Conteúdo<br>
      - CustomFunctions<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-9-requirement-set.md">ExcelApi 1.9</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-10-requirement-set.md">ExcelApi 1.10</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-11-requirement-set.md">ExcelApi 1.11</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-12-requirement-set.md">ExcelApi 1.12</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-13-requirement-set.md">ExcelApi 1.13</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a><br>
      - <a href="../reference/requirement-sets/ribbon-api-requirement-sets.md">RibbonApi 1.1</a><br>
      - <a href="../reference/requirement-sets/shared-runtime-requirement-sets.md">SharedRuntime 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 no Mac<br>(compra avulsa)</td>
    <td>
      - TaskPane<br>
      - Conteúdo<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-2-requirement-set.md">ExcelApi 1.2</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-3-requirement-set.md">ExcelApi 1.3</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-4-requirement-set.md">ExcelApi 1.4</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-5-requirement-set.md">ExcelApi 1.5</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-6-requirement-set.md">ExcelApi 1.6</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-7-requirement-set.md">ExcelApi 1.7</a><br>
      - <a href="../reference/requirement-sets/excel-api-1-8-requirement-set.md">ExcelApi 1.8</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 no Mac<br>(compra avulsa)</td>
    <td>
      - TaskPane<br>
      - Conteúdo </td>
    <td>
      - <a href="../reference/requirement-sets/excel-api-1-1-requirement-set.md">ExcelApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; – Adicionado com atualizações pós-lançamento.*

## <a name="custom-functions-excel-only"></a>Funções personalizadas (somente Excel)

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th>
    <th>Conjuntos de requisitos da API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></th>
  </tr>
  <tr>
    <td>Office na Web</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office no Windows<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
  <tr>
    <td>Office no Mac<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>- CustomFunctions</td>
    <td>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.1</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.2</a><br>
      - <a href="../excel/custom-functions-requirement-sets.md">CustomFunctionsRuntime 1.3</a>
    </td>
    <td></td>
  </tr>
</table>

## <a name="outlook"></a>Outlook

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th>
    <th>Conjuntos de requisitos da API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></th>
  </tr>
  <tr>
    <td>Office na Web<br>(moderno)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Caixa de correio 1.8</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Caixa de correio 1.9</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Caixa de correio 1.10</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office na Web<br>(clássico)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office no Windows<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Módulos</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Caixa de correio 1.8</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md">Caixa de correio 1.9</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.10/outlook-requirement-set-1.10.md">Caixa de correio 1.10</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office 2019 no Windows<br>(compra avulsa)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Módulos</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office 2016 no Windows<br>(compra avulsa)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de suplemento</a><br>
      - <a href="../reference/manifest/extensionpoint.md#module">Módulos</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Caixa de correio 1.4</a><sup>2</sup>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office 2013 no Windows<br>(compra avulsa)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md">Caixa de correio 1.3</a><sup>2</sup><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md">Caixa de correio 1.4</a><sup>2</sup>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office no iOS<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Organizador de compromissos (Redigir): reunião online</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Caixa de correio 1.5</a>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office no Mac<br>(Interface do Usuário atual,<br>conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Caixa de correio 1.8</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office no Mac<br>(nova Interface do Usuário (visualização)<sup>3</sup>,<br>conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7.md">Caixa de correio 1.7</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md">Caixa de correio 1.8</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><sup>1</sup>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office 2019 no Mac<br>(compra avulsa)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Office 2016 no Mac<br>(compra avulsa)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#messagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#messagecomposecommandsurface">Composição da mensagem</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentattendeecommandsurface">Participante do compromisso (Leitura)</a><br>
      - <a href="../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface">Organizador de compromissos (Redigir)</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md"> Caixa de correio 1.5</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6.md">Caixa de correio 1.6</a>
    </td>
    <td>Não disponível</td>
  </tr>
  <tr>
    <td>Outlook no Android<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - <a href="../reference/manifest/extensionpoint.md#mobilemessagereadcommandsurface">Mensagem lida</a><br>
      - <a href="../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface">Organizador de compromissos (Redigir): reunião online</a><br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1.md"> Caixa de correio 1.1</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2.md">Caixa de correio 1.2</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md"> Caixa de correio 1.3</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4.md"> Caixa de correio 1.4</a><br>
      - <a href="../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md">Caixa de correio 1.5</a>
    </td>
    <td>Não disponível</td>
  </tr>
</table>

> [!NOTE]
> <sup>1</sup> Para exigir o conjunto 1.3 da API de Identidade no código do suplemento, verifique se ele tem suporte ligando para `isSetSupported('IdentityAPI', '1.3')`. Não há suporte para declará-lo no manifesto do seu suplemento. Você também pode determinar se a API tem suporte, verificando se ela não é `undefined`. Para mais detalhes, confira [Usar APIs de conjuntos de requisitos posteriores](../reference/requirement-sets/outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).
>
> <sup>2</sup> Adicionado com atualizações pós-lançamento.
>
> <sup>3</sup> O suporte para a nova Interface do Usuário do Mac (visualização) está disponível no Outlook versão 16.38.506. Para mais informações, consulte a seção [Suporte de Suplemento no Outlook na nova Interface do Usuário do Mac](../outlook/compare-outlook-add-in-support-in-outlook-for-mac.md#add-in-support-in-outlook-on-new-mac-ui-preview).

> [!IMPORTANT]
> O suporte ao cliente para um conjunto de requisitos pode ser restringido pelo suporte do servidor Exchange. Consulte [Conjuntos de requisitos da API JavaScript do Outlook](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) para obter detalhes sobre o intervalo de conjuntos de requisitos suportado pelo servidor Exchange e pelos clientes Outlook.

<br/>

## <a name="word"></a>Word

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th>
    <th>Conjuntos de requisitos da API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></th>
  </tr>
  <tr>
    <td>Office na Web</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office no Windows<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 no Windows<br>(compra avulsa)</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 no Windows<br>(compra avulsa)</td>
    <td>- TaskPane</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 no Windows<br>(compra avulsa)</td>
    <td>- TaskPane</td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office no iPad<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>- TaskPane</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office no Mac<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 no Mac<br>(compra avulsa)</td>
    <td>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/word-api-1-2-requirement-set.md">WordApi 1.2</a><br>
      - <a href="../reference/requirement-sets/word-api-1-3-requirement-set.md">WordApi 1.3</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 no Mac<br>(compra avulsa)</td>
    <td>- TaskPane</td>
    <td>
      - <a href="../reference/requirement-sets/word-api-1-1-requirement-set.md">WordApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#bindingevents">BindingEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#customxmlparts">CustomXmlParts</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixbindings">MatrixBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#matrixcoercion">MatrixCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#ooxmlcoercion">OoxmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablebindings">TableBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#tablecoercion">TableCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textbindings">TextBindings</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textfile">TextFile</a>
    </td>
  </tr>
</table>

*&ast; – Adicionado com atualizações pós-lançamento.*

<br/>

## <a name="powerpoint"></a>PowerPoint

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th>
    <th>Conjuntos de requisitos da API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></th>
  </tr>
  <tr>
    <td>Office na Web</td>
    <td>
      - Conteúdo<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office no Windows<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - Conteúdo<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 no Windows<br>(compra avulsa)</td>
    <td>
      - Conteúdo<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 no Windows<br>(compra avulsa)</td>
    <td>
      - Conteúdo<br>
      - TaskPane </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 no Windows<br>(compra avulsa)</td>
    <td>
      - Conteúdo<br>
      - TaskPane </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office no iPad<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - Conteúdo<br>
      - TaskPane </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office no Mac<br>(conectado a uma assinatura do Microsoft 365)</td>
    <td>
      - Conteúdo<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/powerpoint-api-1-1-requirement-set.md">PowerPointApi 1.1</a><br>
      - <a href="../reference/requirement-sets/powerpoint-api-1-2-requirement-set.md">PowerPointApi 1.2</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/identity-api-requirement-sets.md">IdentityAPI 1.3</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-12">ImageCoercion 1.2</a><br>
      - <a href="../reference/requirement-sets/open-browser-window-api-requirement-sets.md">OpenBrowserWindowApi 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2019 no Mac<br>(compra avulsa)</td>
    <td>
      - Conteúdo<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 no Mac<br>(compra avulsa)</td>
    <td>
      - Conteúdo<br>
      - TaskPane </td>
    <td>
       - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a>*<br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#activeview">ActiveView</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#compressedfile">CompressedFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#file">Arquivo</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#pdffile">PdfFile</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

*&ast; – Adicionado com atualizações pós-lançamento.*

<br/>

## <a name="onenote"></a>OneNote

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th>
    <th>Conjuntos de requisitos da API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></th>
  </tr>
  <tr>
    <td>Office na Web</td>
    <td>
      - Conteúdo<br>
      - TaskPane<br>
      - <a href="../reference/requirement-sets/add-in-commands-requirement-sets.md">Comandos de Suplemento</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/onenote-api-requirement-sets.md">OneNoteApi 1.1</a><br>
      - <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a><br>
      - <a href="../reference/requirement-sets/image-coercion-requirement-sets.md#imagecoercion-11">ImageCoercion 1.1</a>
    </td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#documentevents">DocumentEvents</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#htmlcoercion">HtmlCoercion</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#settings">Configurações</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## <a name="project"></a>Project

<table style="width:80%">
  <tr>
    <th>Plataforma</th>
    <th>Pontos de extensão</th>
    <th>Conjuntos de requisitos da API</th>
    <th><a href="../reference/requirement-sets/office-add-in-requirement-sets.md"><b>APIs comuns</b></a></th>
  </tr>
  <tr>
    <td>Office 2019 no Windows<br>(compra avulsa)</td>
    <td>- TaskPane</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2016 no Windows<br>(compra avulsa)</td>
    <td>- TaskPane</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
  <tr>
    <td>Office 2013 no Windows<br>(compra avulsa)</td>
    <td>- TaskPane</td>
    <td>- <a href="../reference/requirement-sets/dialog-api-requirement-sets.md">DialogApi 1.1</a></td>
    <td>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#selection">Selection</a><br>
      - <a href="../reference/requirement-sets/office-add-in-requirement-sets.md#textcoercion">TextCoercion</a>
    </td>
  </tr>
</table>

<br/>

## <a name="see-also"></a>Confira também

- [Visão geral da plataforma Suplementos do Office](office-add-ins.md)
- [Versões do Office e conjuntos de requisitos](../develop/office-versions-and-requirement-sets.md)
- [Conjuntos de requisitos comuns da API](../reference/requirement-sets/office-add-in-requirement-sets.md)
- [Conjuntos de requisitos dos comandos de suplemento](../reference/requirement-sets/add-in-commands-requirement-sets.md)
- [Documentação de Referência da API](../reference/javascript-api-for-office.md)
- [Histórico de atualizações para Microsoft 365 Apps](/officeupdates/update-history-office365-proplus-by-date)
- [Histórico de atualizações do Office 2016 e 2019 (Clique para Executar)](/officeupdates/update-history-office-2019)
- [Histórico de atualizações do Office 2013 (clique para executar)](/officeupdates/update-history-office-2013)
- [Histórico de atualizações do Office 2010, 2013, e 2016 (MSI)](/officeupdates/office-updates-msi)
- [Histórico de atualizações do Outlook 2010, 2013, e 2016 (MSI)](/officeupdates/outlook-updates-msi)
- [Histórico de atualizações do Office para Mac](/officeupdates/update-history-office-for-mac)
- [Desenvolver Suplementos do Office ](../develop/develop-overview.md)
