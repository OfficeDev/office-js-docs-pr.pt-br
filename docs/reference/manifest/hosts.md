---
title: Elemento Hosts no arquivo de manifesto
description: Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 037ac2b5fedbfb1b59b7523382574942fe59a00a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611803"
---
# <a name="hosts-element"></a><span data-ttu-id="be5ca-103">Elemento Hosts</span><span class="sxs-lookup"><span data-stu-id="be5ca-103">Hosts element</span></span>

<span data-ttu-id="be5ca-p101">Especifica o aplicativo cliente do Office no qual o suplemento do Office será ativado. Contém um conjunto de elementos **Host** e suas configurações.</span><span class="sxs-lookup"><span data-stu-id="be5ca-p101">Specifies the Office client application where the Office Add-in will activate. Contains a collection of **Host** elements and their settings.</span></span> 

<span data-ttu-id="be5ca-106">Quando incluído no nó [VersionOverrides](versionoverrides.md), este elemento substitui o elemento **Hosts** na parte pai do manifesto.</span><span class="sxs-lookup"><span data-stu-id="be5ca-106">When included in the [VersionOverrides](versionoverrides.md) node, this element overrides the **Hosts** element in the parent portion of the manifest.</span></span> 

## <a name="child-elements"></a><span data-ttu-id="be5ca-107">Elementos filho</span><span class="sxs-lookup"><span data-stu-id="be5ca-107">Child elements</span></span>

|  <span data-ttu-id="be5ca-108">Elemento</span><span class="sxs-lookup"><span data-stu-id="be5ca-108">Element</span></span> |  <span data-ttu-id="be5ca-109">Obrigatório</span><span class="sxs-lookup"><span data-stu-id="be5ca-109">Required</span></span>  |  <span data-ttu-id="be5ca-110">Descrição</span><span class="sxs-lookup"><span data-stu-id="be5ca-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="be5ca-111">Host</span><span class="sxs-lookup"><span data-stu-id="be5ca-111">Host</span></span>](host.md)    |  <span data-ttu-id="be5ca-112">Sim</span><span class="sxs-lookup"><span data-stu-id="be5ca-112">Yes</span></span>   |  <span data-ttu-id="be5ca-113">Descreve um host e suas configurações.</span><span class="sxs-lookup"><span data-stu-id="be5ca-113">Describes a host and its settings.</span></span> |
