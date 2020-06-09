---
title: Elemento EquivalentAddin no arquivo de manifesto
description: Especifica a compatibilidade COM versões anteriores para um suplemento COM equivalente ou XLL.
ms.date: 06/19/2019
localization_priority: Normal
ms.openlocfilehash: e14fe91bf7a5fe321019acf205ddb1753fedd569
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611558"
---
# <a name="equivalentaddin-element"></a><span data-ttu-id="4d2c2-103">Elemento EquivalentAddin</span><span class="sxs-lookup"><span data-stu-id="4d2c2-103">EquivalentAddin element</span></span>

<span data-ttu-id="4d2c2-104">Especifica a compatibilidade COM versões anteriores para um suplemento COM equivalente ou XLL.</span><span class="sxs-lookup"><span data-stu-id="4d2c2-104">Specifies backwards compatibility for an equivalent COM add-in or XLL.</span></span>

<span data-ttu-id="4d2c2-105">**Tipo de suplemento:** Painel de tarefas, função personalizada</span><span class="sxs-lookup"><span data-stu-id="4d2c2-105">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="4d2c2-106">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="4d2c2-106">Syntax</span></span>

```XML
<EquivalentAddin>
   ...
</EquivalentAddin>
```

## <a name="contained-in"></a><span data-ttu-id="4d2c2-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="4d2c2-107">Contained in</span></span>

[<span data-ttu-id="4d2c2-108">EquivalentAdd-ins</span><span class="sxs-lookup"><span data-stu-id="4d2c2-108">EquivalentAdd-ins</span></span>](equivalentaddins.md)

## <a name="must-contain"></a><span data-ttu-id="4d2c2-109">Deve conter</span><span class="sxs-lookup"><span data-stu-id="4d2c2-109">Must contain</span></span>

[<span data-ttu-id="4d2c2-110">Tipo</span><span class="sxs-lookup"><span data-stu-id="4d2c2-110">Type</span></span>](type.md)

## <a name="can-contain"></a><span data-ttu-id="4d2c2-111">Pode conter</span><span class="sxs-lookup"><span data-stu-id="4d2c2-111">Can contain</span></span>

<span data-ttu-id="4d2c2-112">[ProgID](progid.md) 
 [Nome do arquivo](filename.md)</span><span class="sxs-lookup"><span data-stu-id="4d2c2-112">[ProgId](progid.md)
[FileName](filename.md)</span></span>

## <a name="remarks"></a><span data-ttu-id="4d2c2-113">Comentários</span><span class="sxs-lookup"><span data-stu-id="4d2c2-113">Remarks</span></span>

<span data-ttu-id="4d2c2-114">Para especificar um suplemento de COM como o suplemento equivalente, forneça os `ProgId` `Type` elementos e.</span><span class="sxs-lookup"><span data-stu-id="4d2c2-114">To specify a COM add-in as the equivalent add-in, provide both the `ProgId` and `Type` elements.</span></span> <span data-ttu-id="4d2c2-115">Para especificar um XLL como o suplemento equivalente, forneça os `FileName` `Type` elementos e.</span><span class="sxs-lookup"><span data-stu-id="4d2c2-115">To specify an XLL as the equivalent add-in, provide both the `FileName` and `Type` elements.</span></span>

## <a name="see-also"></a><span data-ttu-id="4d2c2-116">Confira também</span><span class="sxs-lookup"><span data-stu-id="4d2c2-116">See also</span></span>

- [<span data-ttu-id="4d2c2-117">Torne as suas funções personalizadas compatíveis com as funções XLL definidas pelo usuário</span><span class="sxs-lookup"><span data-stu-id="4d2c2-117">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="4d2c2-118">Tornar seu suplemento do Excel compatível com um suplemento de COM existente</span><span class="sxs-lookup"><span data-stu-id="4d2c2-118">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)