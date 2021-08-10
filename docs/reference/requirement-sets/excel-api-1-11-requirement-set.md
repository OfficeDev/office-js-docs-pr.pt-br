---
title: Excel Conjunto de requisitos da API JavaScript 1.11
description: Detalhes sobre o conjunto de requisitos do ExcelApi 1.11.
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 67fb212813608ecb4e72ba5d63952f0228875211d0bf66978b7201fff58c5076
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092651"
---
# <a name="whats-new-in-excel-javascript-api-111"></a>Novidades na EXCEL JavaScript 1.11

O ExcelApi 1.11 aprimorou o suporte para comentários e controles no nível da planilha (como salvar e fechar a planilha). Ele também adicionou acesso às configurações de cultura para ajudar a contabilizar a localização.

| Área de recurso | Descrição | Objetos relevantes |
|:--- |:--- |:--- |
| Menções [de comentário](../../excel/excel-add-ins-comments.md#mentions) |Marca e notifica outros usuários da área de trabalho por meio de comentários. | [Comment](/javascript/api/excel/excel.comment), [CommentRichContent](/javascript/api/excel/excel.commentrichcontent) |
| Resolução de [comentários](../../excel/excel-add-ins-comments.md#resolve-comment-threads) | Resolver threads de comentário e obter o status da resolução. | [Comentário](/javascript/api/excel/excel.comment) |
| [Configurações de cultura](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Obtém configurações do sistema cultural para a caixa de trabalho, como formatação de número. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Aplicativo NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Cortar e colar (moveTo)](../../excel/excel-add-ins-ranges-cut-copy-paste.md) | Replica a funcionalidade de recortar e colar no Excel para um Range. | [Range](/javascript/api/excel/excel.range) |
| [Salvar](../../excel/excel-add-ins-workbooks.md#save-the-workbook) e [Fechar](../../excel/excel-add-ins-workbooks.md#close-the-workbook) a pasta de trabalho | Salve e feche a pasta de trabalho. | [Workbook](/javascript/api/excel/excel.workbook) |
| Eventos de planilha | Eventos adicionais e informações de evento para cálculos de planilha e linhas ocultas. | [WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs), [WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs) |

## <a name="api-list"></a>Lista de API

A tabela a seguir lista as APIs no conjunto de requisitos da API JavaScript Excel 1.11. Para exibir a documentação de referência da API para todas as APIs suportadas pelo Excel conjunto de requisitos da API JavaScript 1.11 ou anterior, consulte Excel APIs no conjunto de requisitos [1.11](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)ou anterior .

| Classe | Campos | Descrição |
|:---|:---|:---|
|[Aplicativo](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureInfo)|Fornece informações com base nas configurações atuais de cultura do sistema.|
||[decimalSeparator](/javascript/api/excel/excel.application#decimalSeparator)|Obtém a cadeia de caracteres usada como separador decimal para valores numéricos.|
||[thousandsSeparator](/javascript/api/excel/excel.application#thousandsSeparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos.|
||[useSystemSeparators](/javascript/api/excel/excel.application#useSystemSeparators)|Especifica se os separadores do sistema de Excel estão habilitados.|
|[Comentário](/javascript/api/excel/excel.comment)|[menções](/javascript/api/excel/excel.comment#mentions)|Obtém as entidades (por exemplo, pessoas) mencionadas nos comentários.|
||[richContent](/javascript/api/excel/excel.comment#richContent)|Obtém o conteúdo rich comment (por exemplo, menções nos comentários).|
||[resolvido](/javascript/api/excel/excel.comment#resolved)|O status do thread de comentário.|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.comment#updateMentions_contentWithMentions_)|Atualiza o conteúdo do comentário com uma cadeia de caracteres especialmente formatada e uma lista de menções.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Cadeia de caracteres de intervalo, conteúdo: cadeia de caracteres \| CommentRichContent, \| contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|Cria um novo comentário com o conteúdo fornecido na célula especificada.|
|[CommentMention](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|O endereço de email da entidade mencionada em um comentário.|
||[id](/javascript/api/excel/excel.commentmention#id)|A ID da entidade.|
||[name](/javascript/api/excel/excel.commentmention#name)|O nome da entidade mencionada em um comentário.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[menções](/javascript/api/excel/excel.commentreply#mentions)|As entidades (por exemplo, pessoas) mencionadas nos comentários.|
||[resolvido](/javascript/api/excel/excel.commentreply#resolved)|O status da resposta ao comentário.|
||[richContent](/javascript/api/excel/excel.commentreply#richContent)|O conteúdo rich comment (por exemplo, menções nos comentários).|
||[updateMentions(contentWithMentions: Excel. CommentRichContent)](/javascript/api/excel/excel.commentreply#updateMentions_contentWithMentions_)|Atualiza o conteúdo do comentário com uma cadeia de caracteres especialmente formatada e uma lista de menções.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: CommentRichContent \| string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|Cria uma resposta de comentário para um comentário.|
|[CommentRichContent](/javascript/api/excel/excel.commentrichcontent)|[menções](/javascript/api/excel/excel.commentrichcontent#mentions)|Uma matriz que contém todas as entidades (por exemplo, pessoas) mencionadas no comentário.|
||[richContent](/javascript/api/excel/excel.commentrichcontent#richContent)|Especifica o conteúdo rico do comentário (por exemplo, conteúdo de comentário com menções, a primeira entidade mencionada tem um atributo ID de 0 e a segunda entidade mencionada tem um atributo ID de 1).|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Obtém o nome da cultura no formato languagecode2-country/regioncode2 (por exemplo, "zh-cn" ou "en-us").|
||[numberFormat](/javascript/api/excel/excel.cultureinfo#numberFormat)|Define o formato culturalmente apropriado de exibição de números.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[numberDecimalSeparator](/javascript/api/excel/excel.numberformatinfo#numberDecimalSeparator)|Obtém a cadeia de caracteres usada como separador decimal para valores numéricos.|
||[numberGroupSeparator](/javascript/api/excel/excel.numberformatinfo#numberGroupSeparator)|Obtém a cadeia de caracteres usada para separar grupos de dígitos à esquerda do decimal para valores numéricos.|
|[Range](/javascript/api/excel/excel.range)|[moveTo(destinationRange: Cadeia de \| caracteres de intervalo)](/javascript/api/excel/excel.range#moveTo_destinationRange_)|Move valores de célula, formatação e fórmulas do intervalo atual para o intervalo de destino, substituindo as informações antigas nessas células.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[adjustIndent(amount: number)](/javascript/api/excel/excel.rangeformat#adjustIndent_amount_)|Ajusta o recuo da formatação do intervalo.|
|[Pasta de trabalho](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close_closeBehavior_)|Fechar a pasta de trabalho atual.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save_saveBehavior_)|Salvar a pasta de trabalho atual.|
|[Planilha](/javascript/api/excel/excel.worksheet)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheet#onRowHiddenChanged)|Ocorre quando o estado oculto de uma ou mais linhas foi alterado em uma planilha específica.|
|[WorksheetCalculatedEventArgs](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|O endereço do intervalo que concluiu o cálculo.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onRowHiddenChanged](/javascript/api/excel/excel.worksheetcollection#onRowHiddenChanged)|Ocorre quando o estado oculto de uma ou mais linhas foi alterado em uma planilha específica.|
|[WorksheetRowHiddenChangedEventArgs](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Obtém o endereço do intervalo que representa a área alterada de uma planilha específica.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changeType)|Obtém o tipo de alteração que representa como o evento foi disparado.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Obtém a origem do evento.|
||[tipo](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Obtém o tipo do evento.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetId)|Obtém a ID da planilha na qual os dados foram alterados.|

## <a name="see-also"></a>Confira também

- [Documentação deReferência da API JavaScript do Excel](/javascript/api/excel?view=excel-js-1.11&preserve-view=true)
- [Conjuntos de requisitos da API JavaScript do Excel](excel-api-requirement-sets.md)
