---
title: Conjunto de requisitos de API JavaScript do Excel 1,11
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1,11.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2ca378364cacab58000607aef15f450e14545ac6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996267"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>O que há de novo na API JavaScript do Excel 1,11

O ExcelApi 1,11 melhorou o suporte para comentários e controles de nível de pasta de trabalho (como salvar e fechar a pasta de trabalho). Ele também adicionou acesso às configurações de cultura para ajudar a sua conta na localização.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Comentários [menciona](../../excel/excel-add-ins-comments.md#mentions) |Marca e notifica outros usuários da pasta de trabalho por meio de comentários. | [Comentário](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| [Resolução](../../excel/excel-add-ins-comments.md#resolve-comment-threads) de comentários | Resolver os threads de comentário e obter o status de resolução. | [Comment](/javascript/api/excel/excel.comment) |
| [Configurações de cultura](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Obtém configurações culturais do sistema para a pasta de trabalho, como formatação de número. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [aplicativo](/javascript/api/excel/excel.application) NumberFormatInfo |
| [Recortar e colar (moveTo)](../../excel/excel-add-ins-ranges-advanced.md#cut-copy-and-paste) | Replica a funcionalidade de recortar e colar no Excel para um intervalo. | [Range](/javascript/api/excel/excel.range) |
| [Salvar](../../excel/excel-add-ins-workbooks.md#save-the-workbook) e [Fechar](../../excel/excel-add-ins-workbooks.md#close-the-workbook) a pasta de trabalho | Salve e feche a pasta de trabalho. | [Workbook](/javascript/api/excel/excel.workbook) |
| Eventos de planilha | Eventos adicionais e informações de eventos para cálculos de planilha e linhas ocultas. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Lista de APIs

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript do Excel 1,11. Para exibir a documentação de referência da API para todas as APIs suportadas pelo conjunto de requisitos de API JavaScript do Excel 1,11 ou anterior, confira [APIs do Excel no conjunto de requisitos 1,11 ou anterior](/javascript/api/excel?view=excel-js-1.11&preserve-view=true).

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Fornece informações com base nas configurações de cultura do sistema atual.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalseparator)|Obtém a cadeia de caracteres usada como o separador decimal para valores numéricos.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsseparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos.|
||[useSystemSeparators](/javascript/api/excel/excel.application#usesystemseparators)|Especifica se os separadores de sistema do Excel estão habilitados.|
|[Comment](/javascript/api/excel/excel.comment)|[menções](/javascript/api/excel/excel.comment#mentions)|Obtém as entidades (por exemplo, pessoas) mencionadas em comentários.|
||[richContent](/javascript/api/excel/excel.comment#richcontent)|Obtém o conteúdo de comentário avançado (por exemplo, menciona em comentários).|
||[Obtido](/javascript/api/excel/excel.comment#resolved)|O status do thread de comentários.|
||[updateMentions (contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Atualiza o conteúdo de comentários com uma cadeia de caracteres especialmente formatada e uma lista de menção.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (cellAddress: \| String de intervalo, Content: \| cadeia de caracteres CommentRichContent, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Cria um novo comentário com o conteúdo fornecido na célula especificada.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|O endereço de email da entidade mencionada em comentário.|
||[id](/javascript/api/excel/excel.commentmention#id)|A ID da entidade.|
||[name](/javascript/api/excel/excel.commentmention#name)|O nome da entidade mencionada em comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[menções](/javascript/api/excel/excel.commentreply#mentions)|As entidades (por exemplo, pessoas) mencionadas em comentários.|
||[Obtido](/javascript/api/excel/excel.commentreply#resolved)|O status de resposta de comentário.|
||[richContent](/javascript/api/excel/excel.commentreply#richcontent)|O conteúdo de comentário avançado (por exemplo, menciona comentários).|
||[updateMentions (contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Atualiza o conteúdo de comentários com uma cadeia de caracteres especialmente formatada e uma lista de menção.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Add (Content: CommentRichContent \| String, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Cria uma resposta de comentário para o comentário.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[menções](/javascript/api/excel/excel.commentrichcontent#mentions)|Uma matriz que contém todas as entidades (por exemplo, pessoas) mencionadas no comentário.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richcontent)|Especifica o conteúdo avançado do comentário (por exemplo, conteúdo de comentários com menção, a primeira entidade mencionada tem um atributo ID 0 e a segunda entidade mencionada tem um atributo ID de 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtém o nome da cultura no formato languagecode2-Country/regioncode2 (por exemplo, "zh-CN" ou "en-US").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberformat)|Define o formato culturalmente apropriado para exibir números.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Obtém a cadeia de caracteres usada como o separador decimal para valores numéricos.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos.|
|[Range](/javascript/api/excel/excel.range)|[moveTo (destinationRange: cadeia de caracteres de intervalo \| )](/javascript/api/excel/excel.range#moveto-destinationrange-)|Move valores de célula, formatação e fórmulas do intervalo atual para o intervalo de destino, substituindo as informações antigas nessas células.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent (valor: número)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Ajusta o recuo da formatação do intervalo.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Fechar a pasta de trabalho atual.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Salvar a pasta de trabalho atual.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Ocorre quando o estado oculto de uma ou mais linhas é alterado em uma planilha específica.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|O endereço do intervalo que concluiu o cálculo.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Ocorre quando o estado oculto de uma ou mais linhas é alterado em uma planilha específica.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Obtém o tipo de alteração que representa como o evento foi acionado.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Obtém o id da planilha na qual os dados são alterados.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
