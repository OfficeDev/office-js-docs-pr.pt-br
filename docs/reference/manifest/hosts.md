---
title: Elemento Hosts no arquivo de manifesto
description: Especifica os aplicativos Office cliente em que o Office o Add-in será ativado.
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9ea6cc9745f47b6e9b1c9bb0232b744304078053
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63341069"
---
# <a name="hosts-element"></a>Elemento Hosts

Especifica os aplicativos Office cliente em que o Office o Add-in será ativado. Contém um conjunto de elementos **Host** e suas configurações. 

## <a name="as-child-of-versionoverrides-element"></a>Como filho do elemento VersionOverrides

As informações nesta seção só se *aplicarão* quando o **elemento Hosts** for filho de [um VersionOverrides](versionoverrides.md).

Esse elemento substitui o **elemento Hosts** no manifesto base.

**Tipo de complemento:** Painel de tarefas, Email

**Válido somente nesses esquemas VersionOverrides**:

- Painel de tarefas 1.0
- Email 1.0
- Email 1.1

Para obter mais informações, consulte [Substituições de versão no manifesto](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="child-elements"></a>Elementos filho

|  Elemento |  Obrigatório  |  Descrição  |
|:-----|:-----|:-----|
|  [Host](host.md)    |  Sim   |  Descreve um host e suas configurações. |
