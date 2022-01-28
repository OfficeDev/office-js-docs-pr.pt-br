---
title: Elemento Requirements no arquivo de manifesto
description: O elemento Requirements especifica o conjunto de requisitos mínimo e os métodos que seu Office Add-in precisa ser ativado pelo Office ou para substituir as configurações de manifesto base.
ms.date: 01/26/2022
ms.localizationpriority: medium
ms.openlocfilehash: e7953ca1e47c492849fe9d0c79384376ffdec347
ms.sourcegitcommit: e837f966d7360ed11b3ff9363ff20380f7d0c45e
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 01/28/2022
ms.locfileid: "62263034"
---
# <a name="requirements-element"></a>Elemento Requirements

O significado desse elemento depende se ele é usado no manifesto [base](#in-the-base-manifest) ou como filho [de **um elemento VersionOverrides**](#as-a-child-of-a-versionoverrides-element).

> [!TIP]
> Antes de usar esse elemento, familiarizar-se com [Especificar Office hosts e requisitos de API](../../develop/specify-office-hosts-and-api-requirements.md)

## <a name="in-the-base-manifest"></a>No manifesto base

Quando usado no manifesto base (ou seja, como filho direto do [OfficeApp](officeapp.md)), o elemento **Requirements** especifica o conjunto mínimo de requisitos de API JavaScript do [Office (conjuntos](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) de requisitos e/ou métodos) que seu Office Add-in precisa ser ativado por Office. O add-in não será ativado em nenhuma combinação de versão e plataforma Office (como Windows, Mac, Web e iOS ou iPad) que não oferece suporte aos métodos e conjuntos de requisitos especificados.

**Tipo de complemento:** Painel de tarefas, Email

## <a name="as-a-child-of-a-versionoverrides-element"></a>Como filho de um elemento VersionOverrides

Quando usado como filho de [VersionOverrides](versionoverrides.md), especifica o conjunto mínimo de requisitos de API JavaScript [do Office](../../develop/office-versions-and-requirement-sets.md#specify-office-applications-and-requirement-sets) (conjuntos de requisitos e/ou métodos) que devem ser suportados pela versão e plataforma do Office (como Windows, Mac, Web e iOS ou iPad) para obter as configurações no elemento **VersionOverrides** que substituem as configurações de manifesto base   para fazer efeito.

Considere um complemento que especifica o requisito A no manifesto base e especifica o requisito B dentro dos **VersionOverrides**. 

- Se a plataforma e Office versão não são suportadas A, o complemento não é ativado e o Office não analisará a seção **VersionOverrides** do manifesto. 
- Se A e B são suportados, o complemento é ativado e toda a marcação no **VersionOverrides** entra em vigor. 
- Se A for suportado, mas B não for, o complemento será ativado e parte da marcação no  **VersionOverrides** entrará em vigor. Especificamente, os elementos filho **dos VersionOverrides** que não substituem os elementos de manifesto base estão em vigor. Por exemplo, um **elemento WebApplicationInfo** ou **equivalentAddins** tem efeito. No entanto, todos os elementos filho **dos VersionOverrides** que substituem um elemento de manifesto base, como **Hosts**, não estão em vigor. Em vez disso, Office usa os valores da marcação de manifesto base que, caso contrário, teriam sido substituídos. 

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associado a esses conjuntos de requisitos**:

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) quando o **VersionOverrides** pai é o tipo Taskpane 1.0.
- [Caixa de correio 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) quando o **VersionOverrides** pai é o tipo Mail 1.0.
- [Caixa de correio 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) quando o **VersionOverrides** pai é o tipo Mail 1.1.

### <a name="remarks"></a>Comentários

O **elemento Requirements** não serve para nada em **um VersionOverrides** se não especificar requisitos adicionais que não sejam especificados em um **Requirements** no manifesto base. Se Office versão e plataforma não suportam os requisitos no manifesto base, o complemento não é ativado e o elemento **VersionOverrides** não é analisado. Por esse motivo, você deve usar um **elemento Requirements** em **um VersionOverrides** somente quando ambas as condições são atendidas:

- Seu complemento tem recursos extras implementados com configuração em **um VersionOverrides** (como Comandos de Complemento) e que exigem um método ou conjunto de requisitos que não é  especificado em um elemento **Requirements** no manifesto base.
- Seu complemento é útil e deve ser ativado (mas sem os recursos extras), mesmo em uma combinação de plataforma e uma versão Office que não oferece suporte aos requisitos necessários para os recursos extras.

> [!TIP]
> Não repita os **elementos De** requisito do manifesto base dentro de **um VersionOverrides**. Fazer isso não tem efeito e é potencialmente enganoso quanto à finalidade do elemento **Requirements** dentro de **um VersionOverrides**.

> [!WARNING]
> Use um grande cuidado antes de usar um elemento **Requirements** em **um VersionOverrides**, porque em combinações de plataforma e versão que não suportam o *requisito, nenhum* dos comandos do add-in será instalado, mesmo aqueles que invocam a funcionalidade que não precisa do *requisito*. Considere, por exemplo, um complemento que tenha dois botões de faixa de opções personalizados. Uma delas chama Office APIs JavaScript disponíveis no conjunto de requisitos **ExcelApi 1.4** (e posterior). As outras CHAMADAS APIs que estão disponíveis apenas no **ExcelApi 1.9** (e posteriores). Se você colocar um requisito para **ExcelApi 1.9** no **VersionOverrides**, quando 1.9 não tiver suporte nenhum botão aparecerá na faixa de opções. Uma estratégia melhor nesse cenário seria usar a técnica descrita em Verificações de tempo de execução [para suporte ao método e ao conjunto de requisitos](../../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). O código invocado pelo segundo botão primeiro usa `isSetSupported` para verificar se há suporte do **ExcelApi 1.9**. Se não for suportado, o código dará ao usuário uma mensagem dizendo que esse recurso do complemento não está disponível em sua versão de Office. 

> [!NOTE]
> Em complementos de email, é possível que um **VersionOverrides** 1.1 seja aninhado dentro de **um VersionOverrides** 1.0. Office sempre usará a versão mais alta **VersionOverrides** que é suportada pela plataforma e Office versão.

## <a name="syntax"></a>Sintaxe

```XML
<Requirements>
   ...
</Requirements>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)
 [VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>Pode conter

|Elemento|Conteúdo|Correio|TaskPane|
|:-----|:-----|:-----|:-----|
|[Sets](sets.md)|x|x|x|
|[Métodos](methods.md)|x||x|

## <a name="see-also"></a>Confira também

Para saber mais sobre os conjuntos de requisitos, confira [Versões do Office e conjuntos de requisitos](../../develop/office-versions-and-requirement-sets.md).
