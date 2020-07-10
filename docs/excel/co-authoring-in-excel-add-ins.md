---
title: Coautoria em suplementos do Excel
description: Saiba como co-autoria de uma pasta de trabalho do Excel armazenada no OneDrive, OneDrive for Business ou SharePoint Online.
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 4414bf64f05c29328c63d0857a6e498495712ff1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093473"
---
# <a name="coauthoring-in-excel-add-ins"></a>Coautoria em suplementos do Excel  

With [coauthoring](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104), multiple people can work together and edit the same Excel workbook simultaneously. All coauthors of a workbook can see another coauthor's changes as soon as that coauthor saves the workbook. To coauthor an Excel workbook, the workbook must be stored in OneDrive, OneDrive for Business, or SharePoint Online.

> [!IMPORTANT]
> No Excel para Microsoft 365, você verá o salvamento automático no canto superior esquerdo. Quando o Salvamento Automático estiver ativado, os coautores verão as respectivas alterações em tempo real. Considere o impacto desse comportamento no design do seu suplemento do Excel. Os usuários podem desativar o Salvamento Automático pelo botão no canto superior esquerdo da janela do Excel.

## <a name="coauthoring-overview"></a>Visão geral da coautoria

Quando você altera o conteúdo de uma pasta de trabalho, o Excel sincroniza automaticamente essas alterações entre todos os coautores. Os coautores podem alterar o conteúdo de uma pasta de trabalho, assim como o código em execução em um suplemento do Excel. Por exemplo, quando o seguinte código JavaScript é executado em um suplemento do Office, o valor de um intervalo é definido como Contoso:

```js
range.values = [['Contoso']];
```
Depois que "Contoso" é sincronizado entre todos os coautores, qualquer usuário ou suplemento em execução na mesma pasta de trabalho verá o novo valor do intervalo.

Coauthoring only synchronizes the content within the shared workbook. Values copied from the workbook to JavaScript variables in an Excel add-in are not synchronized. For example, if your add-in stores the value of a cell (such as 'Contoso') in a JavaScript variable, and then a coauthor changes the value of the cell to 'Example', after synchronization all coauthors see 'Example' in the cell. However, the value of the JavaScript variable is still set to 'Contoso'. Furthermore, when multiple coauthors use the same add-in, each coauthor has their own copy of the variable, which is not synchronized. When you use variables that use workbook content, be sure you check for updated values in the workbook before you use the variable.

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>Usar eventos para gerenciar o estado na memória do suplemento

Excel add-ins can read workbook content (from hidden worksheets and a setting object), and then store it in data structures such as variables. After the original values are copied into any of these data structures, coauthors can update the original workbook content. This means that the copied values in the data structures are now out of sync with the workbook content. When you build your add-ins, be sure to account for this separation of workbook content and values stored in data structures.

For example, you might build a content add-in that displays custom visualizations. The state of your custom visualizations might be saved in a hidden worksheet. When coauthors use the same workbook, the following scenario can occur:

- User A opens the document and the custom visualizations are shown in the workbook. The custom visualizations read data from a hidden worksheet (for example, the color of the visualizations is set to blue).
- User B opens the same document, and starts modifying the custom visualizations. User B sets the color of the custom visualizations to orange. Orange is saved to the hidden worksheet.
- A planilha oculta do Usuário A é atualizada com o novo valor laranja.
- As visualizações personalizadas do Usuário A continuam azuis.

If you want User A's custom visualizations to respond to changes made by coauthors on the hidden worksheet, use the [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) event. This ensures that changes to workbook content made by coauthors is reflected in the state of your add-in.

## <a name="caveats-to-using-events-with-coauthoring"></a>Advertências para usar eventos com coautoria

As described earlier, in some scenarios, triggering events for all coauthors provides an improved user experience. However, be aware that in some scenarios this behavior can produce poor user experiences. 

Por exemplo, em cenários de validação de dados, é comum exibir a interface do usuário em resposta a eventos. O evento [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) descrito na seção anterior é executado quando um usuário local ou coautor (remoto) altera o conteúdo da pasta de trabalho na associação. Se o manipulador de eventos do `BindingDataChanged` evento exibir o UI, os usuários verão a interface do usuário que não está relacionada às alterações em que estavam trabalhando na pasta de trabalho, levando a uma experiência de usuário ruim. Evite a exibição da interface do usuário ao usar eventos no suplemento.

## <a name="see-also"></a>Confira também

- [Sobre a coautoria no Excel (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [Como o Salvamento Automático afeta suplementos e macros (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
