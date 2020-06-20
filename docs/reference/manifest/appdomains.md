---
title: Elemento AppDomains no arquivo de manifesto
description: Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento que seu suplemento do Office usará e deve ser confiável para o Office.
ms.date: 06/12/2020
localization_priority: Normal
ms.openlocfilehash: 751e4ad2ffa5fd50739a855fad48964473b154f1
ms.sourcegitcommit: 9eed5201a3ef556f77ba3b6790f007358188d57d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/17/2020
ms.locfileid: "44778652"
---
# <a name="appdomains-element"></a>Elemento AppDomains

Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento, que o seu suplemento do Office usará e que deve ser confiável para o Office. Isso permite que as páginas nos domínios façam chamadas para Office.js APIs de IFrames no suplemento e têm outros efeitos. Para cada domínio adicional, especifique um elemento **AppDomain**.

 **Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> Há restrições sobre o que pode ser o valor de um elemento **AppDomain** . Para obter mais informações, consulte [AppDomain](appdomain.md).

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

[AppDomain](appdomain.md)

## <a name="remarks"></a>Comentários

Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento [SourceLocation](sourcelocation.md). Esse elemento não pode estar vazio.
