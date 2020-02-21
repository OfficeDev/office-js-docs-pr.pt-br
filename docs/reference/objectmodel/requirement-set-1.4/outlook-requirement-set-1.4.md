---
title: Conjunto de requisitos de API para suplementos do Outlook versão 1.4
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: cb4c8eecd63604aa633ade1a40eb5391b3a62ef2
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165402"
---
# <a name="outlook-add-in-api-requirement-set-14"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.4

O subconjunto de APIs de suplemento do Outlook para as APIs JavaScript do Office inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) que não seja o conjunto de requisitos mais recente.

## <a name="whats-new-in-14"></a>Novidades na versão 1.4

O conjunto de requisitos 1.4 inclui todos os recursos do [Conjunto de requisitos 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). Adicionou acesso ao namespace `Office.ui`.

### <a name="change-log"></a>Log de alterações

- Foi adicionado o [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-): Exibe uma caixa de diálogo em um host do Office.
- Foi adicionado o [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-message-): fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.
- Foi adicionado o objeto [Dialog](/javascript/api/office/office.dialog): o objeto retornado quando o método [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) é chamado.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](../../../quickstarts/outlook-quickstart.md)
- [Conjuntos de requisitos e clientes com suporte](../../requirement-sets/outlook-api-requirement-sets.md)
