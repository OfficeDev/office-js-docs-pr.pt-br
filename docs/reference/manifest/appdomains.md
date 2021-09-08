---
title: Elemento AppDomains no arquivo de manifesto
description: Lista todos os domínios além do domínio especificado no elemento que seu Office Add-in usará e deve ser confiável por `SourceLocation` Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937911"
---
# <a name="appdomains-element"></a>Elemento AppDomains

Lista todos os domínios, além do domínio especificado no elemento, que seu Office Add-in usará e que deve ser confiável por `SourceLocation` Office. Isso permite que as páginas nos domínios façam chamadas Office.js APIs de IFrames dentro do add-in e tenha outros efeitos. Para cada domínio adicional, especifique um elemento **AppDomain**.

 **Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> Há restrições sobre o que pode ser o valor de um **elemento AppDomain.** Para obter mais informações, consulte [AppDomain](appdomain.md).

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

[AppDomain](appdomain.md)

## <a name="remarks"></a>Comentários

Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento [SourceLocation](sourcelocation.md). Esse elemento não pode estar vazio.
