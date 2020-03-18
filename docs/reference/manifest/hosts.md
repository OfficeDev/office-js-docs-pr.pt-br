---
title: Elemento Hosts no arquivo de manifesto
description: Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: cd4e0eecce610b10fdc9dafcde7b807fde425b14
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718101"
---
# <a name="hosts-element"></a><span data-ttu-id="15471-103">Elemento Hosts</span><span class="sxs-lookup"><span data-stu-id="15471-103">Hosts element</span></span>

<span data-ttu-id="15471-p101">Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado. Contém um conjunto de elementos **Host** e suas configurações.</span><span class="sxs-lookup"><span data-stu-id="15471-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="15471-106">Quando incluído no nó [VersionOverrides](versionoverrides.md), este elemento substitui o elemento **Hosts** na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="15471-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="15471-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="15471-107">Child elements</span></span>

|  <span data-ttu-id="15471-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="15471-108">Element</span></span> |  <span data-ttu-id="15471-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="15471-109">Required</span></span>  |  <span data-ttu-id="15471-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="15471-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="15471-111">Host</span><span class="sxs-lookup"><span data-stu-id="15471-111">Host</span></span>](host.md)    |  <span data-ttu-id="15471-112">Sim</span><span class="sxs-lookup"><span data-stu-id="15471-112">Yes</span></span>   |  <span data-ttu-id="15471-113">Descreve um host e suas configurações.</span><span class="sxs-lookup"><span data-stu-id="15471-113">Describes a host and its settings.</span></span> |
