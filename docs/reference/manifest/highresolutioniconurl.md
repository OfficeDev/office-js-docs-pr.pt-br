---
title: Elemento HighResolutionIconUrl no arquivo de manifesto
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: dc8feb92eb8a53351679834a39c012b47f43aad4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432589"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="d101c-102">Elemento HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="d101c-102">HighResolutionIconUrl element</span></span>

<span data-ttu-id="d101c-103">Especifica a URL da imagem que é usada para representar o seu Suplemento do Office na experiência de usuário de inserção e na Office Store em telas de DPI alto.</span><span class="sxs-lookup"><span data-stu-id="d101c-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="d101c-104">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="d101c-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="d101c-105">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="d101c-105">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="d101c-106">Pode conter</span><span class="sxs-lookup"><span data-stu-id="d101c-106">Can contain</span></span>

[<span data-ttu-id="d101c-107">Override</span><span class="sxs-lookup"><span data-stu-id="d101c-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="d101c-108">Atributos</span><span class="sxs-lookup"><span data-stu-id="d101c-108">Attributes</span></span>

|<span data-ttu-id="d101c-109">**Atributo**</span><span class="sxs-lookup"><span data-stu-id="d101c-109">**Attribute**</span></span>|<span data-ttu-id="d101c-110">**Tipo**</span><span class="sxs-lookup"><span data-stu-id="d101c-110">**Type**</span></span>|<span data-ttu-id="d101c-111">**Obrigatório**</span><span class="sxs-lookup"><span data-stu-id="d101c-111">**Required**</span></span>|<span data-ttu-id="d101c-112">**Descrição**</span><span class="sxs-lookup"><span data-stu-id="d101c-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="d101c-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="d101c-113">DefaultValue</span></span>|<span data-ttu-id="d101c-114">cadeia de caracteres (URL)</span><span class="sxs-lookup"><span data-stu-id="d101c-114">string (URL)</span></span>|<span data-ttu-id="d101c-115">obrigatório</span><span class="sxs-lookup"><span data-stu-id="d101c-115">required</span></span>|<span data-ttu-id="d101c-116">Especifica o valor padrão para essa configuração, expresso para a localidade especificada no elemento [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="d101c-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="d101c-117">Observações</span><span class="sxs-lookup"><span data-stu-id="d101c-117">Remarks</span></span>

<span data-ttu-id="d101c-p101">Para um suplemento de email, o ícone é exibido na interface de usuário **Arquivo**  >  **Gerenciar suplementos**. Para um suplemento de conteúdo ou de painel de tarefas, o ícone é exibido na interface de usuário **Inserir**  >  **Suplementos**.</span><span class="sxs-lookup"><span data-stu-id="d101c-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="d101c-120">A imagem deve estar em um dos seguintes formatos: GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="d101c-120">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="d101c-121">Para aplicativos do painel de tarefas e de conteúdo, a resolução de imagem recomendada é 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="d101c-121">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="d101c-122">Para aplicativos de email, a imagem deve ter 128 x 128 pixels.</span><span class="sxs-lookup"><span data-stu-id="d101c-122">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="d101c-123">Para saber mais, confira a seção _Criar uma identidade visual consistente para seu aplicativo_ em [Criar listagens eficazes no AppSource e no Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="d101c-123">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
