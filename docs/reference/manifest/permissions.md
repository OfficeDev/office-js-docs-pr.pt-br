# <a name="permissions-element"></a>Elemento Permissions

Especifica o nível de acesso à API para seu Suplemento do Office. Você deve solicitar permissões com base no princípio do privilégio mínimo.

**Tipo de suplemento:** Conteúdo, Painel de tarefas, Email

## <a name="syntax"></a>Sintaxe

Para suplementos de conteúdo e de painel de tarefas:

```XML
 <Permissions> [Restricted | ReadDocument | ReadAllDocument | WriteDocument | ReadWriteDocument]</Permissions>
```

Para suplementos de e-mail:

```XML
 <Permissions>[Restricted | ReadItem | ReadWriteItem | ReadWriteMailbox]</Permissions>
```

## <a name="contained-in"></a>Contido em

[OfficeApp](officeapp.md)

## <a name="remarks"></a>Comentários

Para saber mais, confira [Solicitar permissões para uso de API em suplementos de conteúdo e de painel de tarefas](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins) e [Entender as permissões de suplemento do Outlook](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).
