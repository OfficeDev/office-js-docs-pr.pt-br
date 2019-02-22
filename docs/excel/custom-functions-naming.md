---
ms.date: 02/08/2019
description: Saiba mais sobre os nomes de funções personalizadas do Excel e evite armadilhas comuns de nomeação.
title: Diretrizes de nomenclatura para funções personalizadas no Excel (visualização)
localization_priority: Normal
ms.openlocfilehash: bdf31879fb6e750fb9dea51f66c55dbc83a2dc90
ms.sourcegitcommit: 8e20e7663be2aaa0f7a5436a965324d171bc667d
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 02/22/2019
ms.locfileid: "30203842"
---
# <a name="naming-guidelines"></a><span data-ttu-id="3b281-103">Diretrizes de nomenclatura</span><span class="sxs-lookup"><span data-stu-id="3b281-103">Naming guidelines</span></span>

<span data-ttu-id="3b281-104">Uma função personalizada é identificada por uma propriedade **ID** e **nome** no arquivo de metadados JSON.</span><span class="sxs-lookup"><span data-stu-id="3b281-104">A custom function is identified by an **id** and **name** property in the JSON metadata file.</span></span> <span data-ttu-id="3b281-105">A ID da função é usada para identificar exclusivamente as funções personalizadas no seu código JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3b281-105">The function id is used to uniquely identify custom functions in your JavaScript code.</span></span> <span data-ttu-id="3b281-106">O nome da função é usado como o nome de exibição que aparece para um usuário no Excel.</span><span class="sxs-lookup"><span data-stu-id="3b281-106">The function name is used as the display name that appears to a user in Excel.</span></span> <span data-ttu-id="3b281-107">Um nome de função pode ser diferente da ID da função, como para fins de localização.</span><span class="sxs-lookup"><span data-stu-id="3b281-107">A function name can differ from the function ID, such as for localization purposes.</span></span> <span data-ttu-id="3b281-108">Mas em geral, ela deve permanecer igual à ID se não houver uma razão convincente para elas diferirem.</span><span class="sxs-lookup"><span data-stu-id="3b281-108">But in general it should stay the same as the ID if there is no compelling reason for them to differ.</span></span>

<span data-ttu-id="3b281-109">Os nomes de função e as IDs de função compartilham alguns requisitos comuns:</span><span class="sxs-lookup"><span data-stu-id="3b281-109">Function names and function IDs share some common requirements:</span></span>

- <span data-ttu-id="3b281-110">Eles devem usar apenas caracteres alfanuméricos (incluindo Unicode), os números de zero a nove, sublinhados e pontos.</span><span class="sxs-lookup"><span data-stu-id="3b281-110">They must only use alphanumeric characters (including Unicode), the numbers zero through nine, underscores, and periods.</span></span>

- <span data-ttu-id="3b281-111">Eles devem começar com uma letra e ter um limite mínimo de três caracteres.</span><span class="sxs-lookup"><span data-stu-id="3b281-111">They must start with a letter and have a minimum limit of three characters.</span></span>

<span data-ttu-id="3b281-112">O `SUM`Excel usa letras maiúsculas para nomes de função internos (como).</span><span class="sxs-lookup"><span data-stu-id="3b281-112">Excel uses uppercase letters for built-in function names (such as `SUM`).</span></span> <span data-ttu-id="3b281-113">Portanto, considere o uso de letras maiúsculas para seus nomes de função personalizada e IDs de função como uma prática recomendada.</span><span class="sxs-lookup"><span data-stu-id="3b281-113">Therefore, consider using uppercase letters for your custom function names and function IDs as a best practice.</span></span>

<span data-ttu-id="3b281-114">Os nomes de função não devem ser nomeados da mesma forma:</span><span class="sxs-lookup"><span data-stu-id="3b281-114">Function names shouldn't be named the same as:</span></span>

- <span data-ttu-id="3b281-115">Qualquer célula entre a1 e XFD1048576 ou qualquer célula entre L1C1 e R1048576C16384.</span><span class="sxs-lookup"><span data-stu-id="3b281-115">Any cells between A1 to XFD1048576 or any cells between R1C1 to R1048576C16384.</span></span>

- <span data-ttu-id="3b281-116">Qualquer função de macro do Excel 4,0 ( `RUN`como `ECHO`,).</span><span class="sxs-lookup"><span data-stu-id="3b281-116">Any Excel 4.0 Macro Function (such as `RUN`, `ECHO`).</span></span>  <span data-ttu-id="3b281-117">Para obter uma lista completa dessas funções, consulte [Este artigo](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span><span class="sxs-lookup"><span data-stu-id="3b281-117">For a full list of these functions, see [this article](https://www.microsoft.com/en-us/download/details.aspx?id=1465).</span></span>

## <a name="naming-conflicts"></a><span data-ttu-id="3b281-118">Conflitos de nomenclatura</span><span class="sxs-lookup"><span data-stu-id="3b281-118">Naming conflicts</span></span>

<span data-ttu-id="3b281-119">Se o nome da função for igual ao nome de uma função em um suplemento que já existe, o **#REF!**</span><span class="sxs-lookup"><span data-stu-id="3b281-119">If your function name is the same as a function name in an add-in that already exists, the **#REF!**</span></span> <span data-ttu-id="3b281-120">o erro aparecerá na sua pasta de trabalho.</span><span class="sxs-lookup"><span data-stu-id="3b281-120">error will appear in your workbook.</span></span>

<span data-ttu-id="3b281-121">Para corrigir um conflito de nomes, altere o nome no suplemento e repita a função.</span><span class="sxs-lookup"><span data-stu-id="3b281-121">To fix a name conflict, change the name in your add-in and try the function again.</span></span> <span data-ttu-id="3b281-122">Você também pode desinstalar o suplemento com o nome conflitante.</span><span class="sxs-lookup"><span data-stu-id="3b281-122">You can also uninstall the add-in with the conflicting name.</span></span> <span data-ttu-id="3b281-123">Ou, se você estiver testando seu suplemento em diferentes ambientes, tente usar um namespace diferente para diferenciar sua função (como NAMESPACE_NAMEOFFUNCTION).</span><span class="sxs-lookup"><span data-stu-id="3b281-123">Or, if you're testing your add-in in different environments, try using a different namespace to differentiate your function (such as NAMESPACE_NAMEOFFUNCTION).</span></span>

<span data-ttu-id="3b281-124">Considere também como você gostaria que as pessoas usem as funções dentro do seu suplemento.</span><span class="sxs-lookup"><span data-stu-id="3b281-124">Also consider how you'd like people to use the functions within your add-in.</span></span> <span data-ttu-id="3b281-125">Em muitos casos, faz sentido adicionar vários argumentos a uma função, em vez de criar várias funções com nomes iguais ou semelhantes.</span><span class="sxs-lookup"><span data-stu-id="3b281-125">In many cases, it makes sense to add multiple arguments to a function rather than create multiple functions with the same or similar names.</span></span>

## <a name="see-also"></a><span data-ttu-id="3b281-126">Confira também</span><span class="sxs-lookup"><span data-stu-id="3b281-126">See also</span></span>

* [<span data-ttu-id="3b281-127">Metadados de funções personalizadas</span><span class="sxs-lookup"><span data-stu-id="3b281-127">Custom functions metadata</span></span>](custom-functions-json.md)
* <span data-ttu-id="3b281-128">[Práticas recomendadas de funções personalizadas](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="3b281-128">[Custom functions best practices](custom-functions-best-practices.md)</span></span>
* [<span data-ttu-id="3b281-129">Tutorial de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="3b281-129">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="3b281-130">Tempo de execução de funções personalizadas do Excel</span><span class="sxs-lookup"><span data-stu-id="3b281-130">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
