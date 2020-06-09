---
title: Elemento AppDomains no arquivo de manifesto
description: Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento que seu suplemento do Office usará para carregar páginas.
ms.date: 07/03/2019
localization_priority: Normal
ms.openlocfilehash: 9183f1815e97bd8d4ac1a7e2cf72d5547d153f7e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608765"
---
# <a name="appdomains-element"></a>Elemento AppDomains

Lista todos os domínios, além do domínio especificado no `SourceLocation` elemento que seu suplemento do Office usará para carregar páginas. Ele também lista os domínios confiáveis dos quais as chamadas de API do Office. js podem ser feitas de IFrames no suplemento. Para cada domínio adicional, especifique um elemento AppDomain.

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
