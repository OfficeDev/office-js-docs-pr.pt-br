# <a name="permissions-element"></a><span data-ttu-id="efc7d-101">Elemento Permissions</span><span class="sxs-lookup"><span data-stu-id="efc7d-101">Permissions element</span></span>

<span data-ttu-id="efc7d-102">Especifica o nível de acesso à API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.</span><span class="sxs-lookup"><span data-stu-id="efc7d-102">Specifies the level of API access for your Office Add-in; you should request permissions based on the principle of least privilege.</span></span>

<span data-ttu-id="efc7d-103">**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email</span><span class="sxs-lookup"><span data-stu-id="efc7d-103">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="efc7d-104">Sintaxe</span><span class="sxs-lookup"><span data-stu-id="efc7d-104">Syntax</span></span>

<span data-ttu-id="efc7d-105">Para suplementos de conteúdo e de painel de tarefas:</span><span class="sxs-lookup"><span data-stu-id="efc7d-105">For content and task pane add-ins:</span></span>

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

<span data-ttu-id="efc7d-106">Para suplementos de e-mail:</span><span class="sxs-lookup"><span data-stu-id="efc7d-106">For mail add-ins:</span></span>

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a><span data-ttu-id="efc7d-107">Contido em</span><span class="sxs-lookup"><span data-stu-id="efc7d-107">Contained in:</span></span>

[<span data-ttu-id="efc7d-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="efc7d-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="efc7d-109">Comentários</span><span class="sxs-lookup"><span data-stu-id="efc7d-109">Remarks</span></span>

<span data-ttu-id="efc7d-110">Para saber mais, confira [Solicitar permissões para uso de API em suplementos de conteúdo e de painel de tarefas](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [Entender as permissões de suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="efc7d-110">For more detail, see [Requesting permissions for API use in content and task pane add-ins](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) and [Understanding Outlook add-in permissions](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>
