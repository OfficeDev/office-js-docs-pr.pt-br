---
title: Coautoria em suplementos do Excel
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d0ef481f9efd6b977963091d5d0a123c30a5a789
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="coauthoring-in-excel-add-ins"></a>Coautoria em suplementos do Excel  

Com a [coautoria](https://support.office.com/en-US/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104), v?rias pessoas podem trabalhar juntas e editar simultaneamente a mesma pasta de trabalho do Excel. Todos os coautores de uma pasta de trabalho podem ver as altera??es de outros coautores assim que o coautor salva a pasta de trabalho. Para ser coautor de uma pasta de trabalho do Excel, esta deve ser armazenada no OneDrive, OneDrive for Business ou SharePoint Online.

> [!IMPORTANT]
> No Excel 2016 para Office 365, voc? ver? o Salvamento Autom?tico no canto superior esquerdo. Quando o Salvamento Autom?tico estiver ativado, os coautores ver?o as respectivas altera??es em tempo real. Considere o impacto desse comportamento no design do seu suplemento do Excel. Os usu?rios podem desativar o Salvamento Autom?tico pelo bot?o no canto superior esquerdo da janela do Excel.

A coautoria est? dispon?vel nas seguintes plataformas:

- Excel Online
- Excel para Android
- Excel para iOS
- Excel Mobile para Windows 10
- Excel para Windows Desktop para clientes do Office 365 (compila??o 16.0.8326.2076 ou posterior do Windows Desktop, que est? dispon?vel para clientes do canal atual em vigor desde agosto de 2017)

## <a name="coauthoring-overview"></a>Vis?o geral da coautoria
 
Quando voc? altera o conte?do de uma pasta de trabalho, o Excel sincroniza automaticamente essas altera??es entre todos os coautores. Os coautores podem alterar o conte?do de uma pasta de trabalho, assim como o c?digo em execu??o em um suplemento do Excel. Por exemplo, quando o seguinte c?digo JavaScript ? executado em um suplemento do Office, o valor de um intervalo ? definido como Contoso:

```js
range.values = [['Contoso']];
```
Depois que "Contoso" ? sincronizado entre todos os coautores, qualquer usu?rio ou suplemento em execu??o na mesma pasta de trabalho ver? o novo valor do intervalo. 

A coautoria sincroniza apenas o conte?do dentro da pasta de trabalho compartilhada. Os valores copiados da pasta de trabalho em vari?veis de JavaScript em um suplemento do Excel n?o s?o sincronizados. Por exemplo, se seu suplemento armazenar o valor de uma c?lula (como "Contoso") em uma vari?vel de JavaScript e um coautor alterar o valor da c?lula para "Exemplo", ap?s a sincroniza??o todos os coautores ver?o "Exemplo" na c?lula. No entanto, o valor da vari?vel de JavaScript continuar? definido como "Contoso". Al?m disso, quando v?rios autores usarem o mesmo suplemento, cada coautor ter? sua pr?pria c?pia da vari?vel, que n?o ? sincronizada. Quando voc? usar vari?veis que usam o conte?do da pasta de trabalho, n?o se esque?a de verificar se h? valores atualizados na pasta de trabalho antes de usar a vari?vel. 

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>Usar eventos para gerenciar o estado na mem?ria do suplemento
 
Os suplementos do Excel podem ler conte?do da pasta de trabalho (de planilhas ocultas e um objeto de configura??o) e armazen?-lo em estruturas de dados, como vari?veis. Depois que os valores originais s?o copiados em qualquer uma dessas estruturas de dados, os coautores podem atualizar o conte?do da pasta de trabalho original. Isso significa que os valores copiados nas estruturas de dados agora est?o fora de sincronia com o conte?do da pasta de trabalho. Ao criar seus suplementos, lembre-se dessa separa??o do conte?do da pasta de trabalho e dos valores armazenados em estruturas de dados.

Por exemplo, voc? pode criar um suplemento de conte?do que exibe visualiza??es personalizadas. O estado de suas visualiza??es personalizadas pode ser salvo em uma planilha oculta. Quando coautores usarem a mesma pasta de trabalho, o seguinte cen?rio poder? ocorrer:

- O Usu?rio A abre o documento e as visualiza??es personalizadas s?o mostradas na pasta de trabalho. As visualiza??es personalizadas leem dados de uma planilha oculta (por exemplo, a cor das visualiza??es ? definida como azul).
- O usu?rio B abre o mesmo documento e come?a a modificar as visualiza??es personalizadas. O usu?rio B define a cor das visualiza??es personalizadas para laranja. A cor laranja ? salva para a planilha oculta.
- A planilha oculta do Usu?rio A ? atualizada com o novo valor laranja.
- As visualiza??es personalizadas do Usu?rio A continuam azuis. 

Se quiser que as visualiza??es personalizadas do Usu?rio A respondam ?s altera??es feitas pelos coautores na planilha oculta, use o evento [BindingDataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent). Isso garante que as altera??es no conte?do da pasta de trabalho feitas pelos coautores sejam refletidas no estado do seu suplemento.

## <a name="caveats-to-using-events-with-coauthoring"></a>Advert?ncias para usar eventos com coautoria 

Conforme descrito anteriormente, em alguns cen?rios, acionar eventos para todos os coautores proporciona uma experi?ncia do usu?rios aprimorada. No entanto, lembre-se de que, em alguns cen?rios, esse comportamento pode resultar em uma m? experi?ncia do usu?rio. 

Por exemplo, em cen?rios de valida??o de dados, ? comum exibir a interface do usu?rio em resposta a eventos. O evento [BindingDataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) descrito na se??o anterior ? executado quando um usu?rio local ou coautor (remoto) altera o conte?do da pasta de trabalho na associa??o. Se o manipulador de eventos do evento **BindingDataChanged** exibir a interface do usu?rio, os usu?rios ver?o a interface do usu?rio que n?o est? relacionada ?s altera??es em que eles estavam trabalhando na pasta de trabalho, levando a uma m? experi?ncia do usu?rio. Evite a exibi??o da interface do usu?rio ao usar eventos no suplemento.

## <a name="see-also"></a>Veja tamb?m 

- [Sobre a coautoria no Excel (VBA)](https://msdn.microsoft.com/en-us/vba/excel-vba/articles/about-coauthoring-in-excel) 
- [Como o Salvamento Autom?tico afeta suplementos e macros (VBA)](https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/how-autosave-impacts-addins-and-macros) 
