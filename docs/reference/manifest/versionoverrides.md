---
title: Elemento VersionOverrides no arquivo de manifesto
description: Documentação de referência do elemento VersionOverrides para Office arquivos XML (manifesto de complementos).
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 91fdaad1bc92ee7baa0b7c2b05aefecf994a93fa
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744972"
---
# <a name="versionoverrides-element"></a>Elemento VersionOverrides

Esse elemento contém informações para recursos que não são suportados no manifesto base. Sua marcação filho pode substituir parte da marcação no manifesto base (ou em **um VersionOverrides pai**). **VersionOverrides** é um elemento filho do elemento [raiz do OfficeApp](officeapp.md) no manifesto ou de **um elemento VersionOverrides** pai. Esse elemento é suportado no esquema de manifesto v1.1 e posterior, mas é definido em esquemas versionOverrides separados.

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="attributes"></a>Atributos

|  Atributo  |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  **xmlns**       |  Sim  |  O namespace de esquema VersionOverrides. Os valores permitidos variam dependendo `<VersionOverrides>` do valor **xsi:type** deste elemento e do **valor xsi:type** do elemento `<OfficeApp>` pai. Consulte [Valores de namespace abaixo](#namespace-values) .|
|  **xsi:type**  |  Sim  | A versão do esquema. Nesse momento, os únicos valores válidos são `VersionOverridesV1_0` e `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Valores de namespace

O seguinte lista o valor necessário do **atributo xmlns** , dependendo do **valor xsi:type** do elemento raiz `<OfficeApp>` .

- **TaskPaneApp dá** suporte apenas à versão 1.0 de VersionOverrides, e **os xmlns** devem ser `http://schemas.microsoft.com/office/taskpaneappversionoverrides`.
- **ContentApp** dá suporte apenas à versão 1.0 de VersionOverrides, e os **xmlns** devem ser `http://schemas.microsoft.com/office/contentappversionoverrides`.
- **MailApp** dá suporte às versões 1.0 e 1.1 de VersionOverrides, portanto, o valor de **xmlns** varia dependendo `<VersionOverrides>` do valor **xsi:type** deste elemento:
  - Quando **xsi:type** for `VersionOverridesV1_0`, **xmlns** devem ser `http://schemas.microsoft.com/office/mailappversionoverrides`.
  - Quando **xsi:type** for `VersionOverridesV1_1`, **xmlns** devem ser `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.

> [!NOTE]
> Atualmente, somente Outlook 2016 ou posterior suporta o esquema VersionOverrides v1.1 e o `VersionOverridesV1_1` tipo.

## <a name="variant-schemas"></a>Esquemas variantes

Há um esquema diferente para cada um dos valores **xmlns** possíveis, portanto, cada um tem uma página de referência separada.

- [Painel de Tarefas versionOverrides 1.0](versionoverrides-1-0-taskpane.md)
- [Conteúdo do VersionOverrides 1.0](versionoverrides-1-0-content.md)
- [Email do VersionOverrides 1.0](versionoverrides-1-0-mail.md)
- [Email do VersionOverrides 1.1](versionoverrides-1-1-mail.md)
