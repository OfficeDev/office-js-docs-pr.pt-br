---
title: Coautoria em suplementos do Excel
description: Aprenda a coautor uma Excel de trabalho armazenada em OneDrive, OneDrive for Business ou SharePoint Online.
ms.date: 07/23/2020
localization_priority: Normal
ms.openlocfilehash: 6901df106f9b70bc0265d78757ec7416380fc76c
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671482"
---
# <a name="coauthoring-in-excel-add-ins"></a>Coautoria em suplementos do Excel  

Com a [coautoria](https://support.office.com/article/Collaborate-on-Excel-workbooks-at-the-same-time-with-co-authoring-7152aa8b-b791-414c-a3bb-3024e46fb104), várias pessoas podem trabalhar juntas e editar simultaneamente a mesma pasta de trabalho do Excel. Todos os coautores de uma pasta de trabalho podem ver as alterações de outros coautores assim que o coautor salva a pasta de trabalho. Para ser coautor de uma pasta de trabalho do Excel, esta deve ser armazenada no OneDrive, OneDrive for Business ou SharePoint Online.

> [!IMPORTANT]
> Em Excel para Microsoft 365, você notará o AutoSave no canto superior esquerdo. Quando o Salvamento Automático estiver ativado, os coautores verão as respectivas alterações em tempo real. Considere o impacto desse comportamento no design do seu suplemento do Excel. Os usuários podem desativar o Salvamento Automático pelo botão no canto superior esquerdo da janela do Excel.

## <a name="coauthoring-overview"></a>Visão geral da coautoria

Quando você altera o conteúdo de uma pasta de trabalho, o Excel sincroniza automaticamente essas alterações entre todos os coautores. Os coautores podem alterar o conteúdo de uma pasta de trabalho, assim como o código em execução em um suplemento do Excel. Por exemplo, quando o código JavaScript a seguir é executado em um Office, o valor de um intervalo é definido como Contoso.

```js
range.values = [['Contoso']];
```

Depois que "Contoso" é sincronizado entre todos os coautores, qualquer usuário ou suplemento em execução na mesma pasta de trabalho verá o novo valor do intervalo.

A coautoria sincroniza apenas o conteúdo dentro da pasta de trabalho compartilhada. Os valores copiados da pasta de trabalho em variáveis de JavaScript em um suplemento do Excel não são sincronizados. Por exemplo, se seu suplemento armazenar o valor de uma célula (como "Contoso") em uma variável de JavaScript e um coautor alterar o valor da célula para "Exemplo", após a sincronização todos os coautores verão "Exemplo" na célula. No entanto, o valor da variável de JavaScript continuará definido como "Contoso". Além disso, quando vários autores usarem o mesmo suplemento, cada coautor terá sua própria cópia da variável, que não é sincronizada. Quando você usar variáveis que usam o conteúdo da pasta de trabalho, não se esqueça de verificar se há valores atualizados na pasta de trabalho antes de usar a variável.

## <a name="use-events-to-manage-the-in-memory-state-of-your-add-in"></a>Usar eventos para gerenciar o estado na memória do suplemento

Os suplementos do Excel podem ler conteúdo da pasta de trabalho (de planilhas ocultas e um objeto de configuração) e armazená-lo em estruturas de dados, como variáveis. Depois que os valores originais são copiados em qualquer uma dessas estruturas de dados, os coautores podem atualizar o conteúdo da pasta de trabalho original. Isso significa que os valores copiados nas estruturas de dados agora estão fora de sincronia com o conteúdo da pasta de trabalho. Ao criar seus suplementos, lembre-se dessa separação do conteúdo da pasta de trabalho e dos valores armazenados em estruturas de dados.

Por exemplo, você pode criar um suplemento de conteúdo que exibe visualizações personalizadas. O estado de suas visualizações personalizadas pode ser salvo em uma planilha oculta. Quando os coautores usam a mesma lista de trabalho, o cenário a seguir pode ocorrer.

- O Usuário A abre o documento e as visualizações personalizadas são mostradas na pasta de trabalho. As visualizações personalizadas leem dados de uma planilha oculta (por exemplo, a cor das visualizações é definida como azul).
- O usuário B abre o mesmo documento e começa a modificar as visualizações personalizadas. O usuário B define a cor das visualizações personalizadas para laranja. A cor laranja é salva para a planilha oculta.
- A planilha oculta do Usuário A é atualizada com o novo valor laranja.
- As visualizações personalizadas do Usuário A continuam azuis.

Se quiser que as visualizações personalizadas do Usuário A respondam às alterações feitas pelos coautores na planilha oculta, use o evento [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs). Isso garante que as alterações no conteúdo da pasta de trabalho feitas pelos coautores sejam refletidas no estado do seu suplemento.

## <a name="caveats-to-using-events-with-coauthoring"></a>Advertências para usar eventos com coautoria

Conforme descrito anteriormente, em alguns cenários, acionar eventos para todos os coautores proporciona uma experiência do usuários aprimorada. No entanto, lembre-se de que, em alguns cenários, esse comportamento pode resultar em uma má experiência do usuário.

Por exemplo, em cenários de validação de dados, é comum exibir a interface do usuário em resposta a eventos. O evento [BindingDataChanged](/javascript/api/office/office.bindingdatachangedeventargs) descrito na seção anterior é executado quando um usuário local ou coautor (remoto) altera o conteúdo da pasta de trabalho na associação. Se o manipulador de eventos do evento exibir a interface do usuário, os usuários verão a interface do usuário que não está relacionada às alterações em que estavam trabalhando na guia de trabalho, levando a uma experiência de usuário `BindingDataChanged` ruim. Evite a exibição da interface do usuário ao usar eventos no suplemento.

## <a name="avoiding-table-row-coauthoring-conflicts"></a>Evitando conflitos de coautor da linha de tabela

É um problema conhecido que as chamadas para a [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add_index__values_) API podem causar conflitos de coautoria. Não recomendamos o uso dessa API se você antecipar que seu complemento será executado enquanto outros usuários estão editando a workbook do complemento (especificamente, se eles estão editando a tabela ou qualquer intervalo sob a tabela). As orientações a seguir devem ajudá-lo a evitar problemas com o método (e evitar a acionamento da barra amarela Excel mostra que solicita que os usuários `TableRowCollection.add` atualizem).

1. Use [`Range.values`](/javascript/api/excel/excel.range#values) em vez de [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add_index__values_) . Definir os `Range` valores diretamente abaixo da tabela expande automaticamente a tabela. Caso contrário, adicionar linhas de tabela por meio das APIs resulta em conflitos `Table` de mesclagem para usuários de coauth.
1. Não deve haver regras [de validação de](https://support.microsoft.com/office/apply-data-validation-to-cells-29fecbcc-d1b9-42c1-9d76-eff3ce5f7249) dados aplicadas a células abaixo da tabela, a menos que a validação de dados seja aplicada a toda a coluna.
1. Se houver dados sob a tabela, o complemento precisará lidar com isso antes de definir o valor do intervalo. Usar [`Range.insert`](/javascript/api/excel/excel.range#insert_shift_) para inserir uma linha vazia moverá os dados e dará espaço para a tabela de expansão. Caso contrário, você corre o risco de sobrescrever células abaixo da tabela.
1. Não é possível adicionar uma linha vazia a uma tabela com `Range.values` . A tabela só se expande automaticamente se os dados estão presentes nas células diretamente abaixo da tabela. Use dados temporários ou colunas ocultas como solução alternativa para adicionar uma linha de tabela vazia.

## <a name="see-also"></a>Confira também

- [Sobre a coautoria no Excel (VBA)](/office/vba/excel/concepts/about-coauthoring-in-excel)
- [Como o Salvamento Automático afeta suplementos e macros (VBA)](/office/vba/library-reference/concepts/how-autosave-impacts-addins-and-macros)
