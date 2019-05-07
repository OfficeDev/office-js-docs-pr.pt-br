---
ms.date: 05/03/2019
description: Saiba mais sobre os nomes de funções personalizadas do Excel e evite armadilhas comuns de nomeação.
title: Diretrizes de nomenclatura para funções personalizadas no Excel
localization_priority: Normal
ms.openlocfilehash: 3abe04eebfa703666b70ecbde1c68ab0c942003c
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628043"
---
# <a name="naming-guidelines"></a><span data-ttu-id="be69a-103">Diretrizes de nomenclatura</span><span class="sxs-lookup"><span data-stu-id="be69a-103">Naming guidelines</span></span>

<span data-ttu-id="be69a-104">Uma função personalizada é identificada por uma propriedade **ID** e **nome** no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="be69a-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

- <span data-ttu-id="be69a-105">A função `id` é usada para identificar exclusivamente as funções personalizadas no seu código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="be69a-105">The function `id` is used to uniquely identify custom functions in your JavaScript code.</span></span> 
- <span data-ttu-id="be69a-106">A função `name` é usada como o nome de exibição que aparece para um usuário no Excel.</span><span class="sxs-lookup"><span data-stu-id="be69a-106">The function `name` is used as the display name that appears to a user in Excel.</span></span> 

<span data-ttu-id="be69a-107">Uma função `name` pode ser diferente da função `id`, como para fins de localização.</span><span class="sxs-lookup"><span data-stu-id="be69a-107">A function `name` can differ from the function `id`, such as for localization purposes.</span></span> <span data-ttu-id="be69a-108">Em geral, uma função `name` deve permanecer igual a `id` se não houver um motivo convincente para elas diferirem.</span><span class="sxs-lookup"><span data-stu-id="be69a-108">In general, a function's `name` should stay the same as the `id` if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="be69a-109">Uma função `name` e `id` compartilhar alguns requisitos comuns:</span><span class="sxs-lookup"><span data-stu-id="be69a-109">A function's `name` and `id` share some common requirements:</span></span>

- <span data-ttu-id="be69a-110">Uma função `id` pode usar apenas caracteres de A a Z, números de zero a nove, sublinhados e pontos.</span><span class="sxs-lookup"><span data-stu-id="be69a-110">A function's `id` may only use characters A through Z, numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="be69a-111">Uma função `name` pode usar caracteres alfabéticos Unicode, sublinhados e pontos.</span><span class="sxs-lookup"><span data-stu-id="be69a-111">A function's `name` may use any Unicode alphabetic characters, underscores, and periods.</span></span>

- <span data-ttu-id="be69a-112">Ambas funcionam `name` e `id` devem começar com uma letra e ter um limite mínimo de três caracteres.</span><span class="sxs-lookup"><span data-stu-id="be69a-112">Both function `name` and `id` must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="be69a-113">O `SUM`Excel usa letras maiúsculas para nomes de função internos (como).</span><span class="sxs-lookup"><span data-stu-id="be69a-113">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="be69a-114">Portanto, considere o uso de letras maiúsculas para a `name` função `id` personalizada e como uma prática recomendada.</span><span class="sxs-lookup"><span data-stu-id="be69a-114">Therefore, consider using uppercase letters for your custom function's `name` and `id` as a best practice.</span></span>

<span data-ttu-id="be69a-115">Uma função não `name` deve ser nomeada da mesma forma:</span><span class="sxs-lookup"><span data-stu-id="be69a-115">A function's `name` shouldn't be named the same as:</span></span>

- <span data-ttu-id="be69a-116">Qualquer célula entre a1 e XFD1048576 ou qualquer célula entre L1C1 e R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="be69a-116">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="be69a-117">Qualquer função de macro do Excel 4,0 ( `RUN`como `ECHO`,).</span><span class="sxs-lookup"><span data-stu-id="be69a-117">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="be69a-118">Para obter uma lista completa dessas funções, consulte [Este artigo](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span><span class="sxs-lookup"><span data-stu-id="be69a-118">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="be69a-119">Conflitos de nomenclatura</span><span class="sxs-lookup"><span data-stu-id="be69a-119">Naming conflicts</span></span>

<span data-ttu-id="be69a-120">Se sua função `name` for igual a uma função `name` em um suplemento que já existe, o **#REF!**</span><span class="sxs-lookup"><span data-stu-id="be69a-120">If your function `name` is the same as a function `name` in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="be69a-121">o erro aparecerá na sua pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="be69a-121">error will appear in your workbook.</span></span>

<span data-ttu-id="be69a-122">Para corrigir um conflito de nomenclatura, altere `name` o em seu suplemento e repita a função.</span><span class="sxs-lookup"><span data-stu-id="be69a-122">To fix a naming conflict, change the `name` in your add-in and try the function again.</span></span> <span data-ttu-id="be69a-123">Você também pode desinstalar o suplemento com o nome conflitante.</span><span class="sxs-lookup"><span data-stu-id="be69a-123">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="be69a-124">Ou, se você estiver testando seu suplemento em diferentes ambientes, tente usar um namespace diferente para diferenciar sua função (como `NAMESPACE_NAMEOFFUNCTION`).</span><span class="sxs-lookup"><span data-stu-id="be69a-124">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as `NAMESPACE_NAMEOFFUNCTION`).</span></span>

## <a name="best-practices"></a><span data-ttu-id="be69a-125">Práticas recomendadas</span><span class="sxs-lookup"><span data-stu-id="be69a-125">Best practices</span></span>

- <span data-ttu-id="be69a-126">Considere adicionar vários argumentos a uma função em vez de criar várias funções com nomes iguais ou semelhantes.</span><span class="sxs-lookup"><span data-stu-id="be69a-126">Consider adding multiple arguments to a function rather than creating multiple functions with the same or similar names.</span></span>
- <span data-ttu-id="be69a-127">Os nomes de função devem indicar a ação da função, como `=GETZIPCODE` em vez `ZIPCODE`de.</span><span class="sxs-lookup"><span data-stu-id="be69a-127">Function names should indicate the action of the function, such as `=GETZIPCODE` instead of `ZIPCODE`.</span></span>
- <span data-ttu-id="be69a-128">Evite abreviações ambíguas em nomes de funções.</span><span class="sxs-lookup"><span data-stu-id="be69a-128">Avoid ambiguous abbreviations in function names.</span></span> <span data-ttu-id="be69a-129">A clareza é mais importante do que a brevidade.</span><span class="sxs-lookup"><span data-stu-id="be69a-129">Clarity is more important than brevity.</span></span> <span data-ttu-id="be69a-130">Escolha um nome como `=INCREASETIME` em vez `=INC`de.</span><span class="sxs-lookup"><span data-stu-id="be69a-130">Choose a name like `=INCREASETIME` rather than `=INC`.</span></span>
- <span data-ttu-id="be69a-131">Use consistentemente os mesmos verbos para funções que executam ações semelhantes.</span><span class="sxs-lookup"><span data-stu-id="be69a-131">Consistently use the same verbs for functions which perform similar actions.</span></span> <span data-ttu-id="be69a-132">Por exemplo, use `=DELETEZIPCODE` e `=DELETEADDRESS`, em vez `=DELETEZIPCODE` de `=REMOVEADDRESS`e.</span><span class="sxs-lookup"><span data-stu-id="be69a-132">For example, use `=DELETEZIPCODE` and `=DELETEADDRESS`, rather than `=DELETEZIPCODE` and `=REMOVEADDRESS`.</span></span>

## <a name="localizing-function-names"></a><span data-ttu-id="be69a-133">Localizando nomes de função</span><span class="sxs-lookup"><span data-stu-id="be69a-133">Localizing function names</span></span>

<span data-ttu-id="be69a-134">Você pode localizar seus nomes de função para idiomas diferentes usando arquivos JSON separados e substituir valores no arquivo de manifesto do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="be69a-134">You can localize your function names for different languages using separate JSON files and override values in your add-in's manifest file.</span></span> <span data-ttu-id="be69a-135">Como prática recomendada, evite dar às funções uma `id` ou `name` que é uma função interna do Excel em outro idioma, pois isso pode causar conflito com funções localizadas.</span><span class="sxs-lookup"><span data-stu-id="be69a-135">As a best practice, avoid giving your functions an `id` or `name` that is a built-in Excel function in another language as this could conflict with localized functions.</span></span>

<span data-ttu-id="be69a-136">Para obter informações completas sobre a localização, consulte [localizar funções personalizadas](custom-functions-localize.md)</span><span class="sxs-lookup"><span data-stu-id="be69a-136">For full information on localizing, see [Localize custom functions](custom-functions-localize.md)</span></span>

## <a name="next-steps"></a><span data-ttu-id="be69a-137">Próximas etapas</span><span class="sxs-lookup"><span data-stu-id="be69a-137">Next steps</span></span>
<span data-ttu-id="be69a-138">Saiba mais sobre [as práticas recomendadas de tratamento de erros](custom-functions-errors.md).</span><span class="sxs-lookup"><span data-stu-id="be69a-138">Learn about [error handling best practices](custom-functions-errors.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="be69a-139">Confira também</span><span class="sxs-lookup"><span data-stu-id="be69a-139">See also</span></span>

* [<span data-ttu-id="be69a-140">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="be69a-140">Custom functions metadata</span></span>](custom-functions-json.md)
* <span data-ttu-id="be69a-141">[Práticas recomendadas para as funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="be69a-141">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="be69a-142">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="be69a-142">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="be69a-143">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="be69a-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
