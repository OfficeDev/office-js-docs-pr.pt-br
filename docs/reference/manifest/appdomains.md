---
title: Elemento AppDomains no arquivo de manifesto
description: ''
ms.date: 12/13/2018
localization_priority: Normal
ms.openlocfilehash: 65391c9529e7ddaa9726d0b58accf90c5b9babef
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450643"
---
# <a name="appdomains-element"></a>Elemento AppDomains

Lista qualquer domínio além do domínio especificado no elemento SourceLocation que seu Suplemento do Office utilizará para carregar páginas. Para cada domínio adicional, especifique um elemento AppDomain.

 **Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

```XML
<AppDomains>
    <AppDomain>AppDomain1</AppDomain>
    <AppDomain>AppDomain2</AppDomain>
</AppDomains>
```

> [!IMPORTANT]
> O valor de cada elemento **AppDomain** deve incluir o protocolo (por exemplo, `<AppDomain>https://myappdomain<AppDomain>`).

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="can-contain"></a>Pode conter

[AppDomain](appdomain.md)

## <a name="remarks"></a>Comentários

Por padrão, o seu suplemento pode carregar qualquer página que esteja no mesmo domínio que o local especificado no elemento [SourceLocation](sourcelocation.md). Para carregar páginas que não estejam no mesmo domínio do que o suplemento, especifique os domínios usando os elementos **AppDomains** e **AppDomain**. Esse elemento não pode estar vazio.
