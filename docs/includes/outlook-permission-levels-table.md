|**Nome canônico do</br>nível de permissão**|**Nome do manifesto XML**|**Nome do manifesto do Teams**|**Descrição do resumo**|
|:-----|:-----|:-----|:-----|
|**Restrito**|Restricted|MailboxItem.Restricted.User|Permite o uso de entidades, mas não expressões regulares. |
|**ler item**|ReadItem|MailboxItem.Read.User|Além do que é permitido **em restrito**, ele permite:<ul><li>expressões regulares</li><li>acesso de leitura para a API do suplemento do Outlook</li><li>obter as propriedades do item e o token de retorno de chamada</li></ul> |
|**item de leitura/gravação**|ReadWriteItem|MailboxItem.ReadWrite.User|Além do que é permitido no **item de leitura**, ele permite:<ul><li>acesso completo à API do Suplemento do Outlook, exceto `makeEwsRequestAsync`</li><li>definição das propriedades do item</li></ul> |
|**caixa de correio de leitura/gravação**|ReadWriteMailbox|Mailbox.ReadWrite.User|Além do que é permitido no **item de leitura/** gravação, ele permite:<ul><li>criar, ler, gravar itens e pastas</li><li>enviar itens</li><li>chamar [makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)</li></ul> |

As permissões são declaradas no manifesto. A marcação varia dependendo do tipo de manifesto.

- **Manifesto XML**: use o **\<Permissions\>** elemento.
- **Manifesto do Teams (versão prévia):** use a propriedade "name" de um objeto na matriz "authorization.permissions.resourceSpecific".

> [!NOTE]
>
> - Há uma permissão complementar necessária para suplementos que usam o recurso de acréscimo ao enviar. Com o manifesto XML, você especifica a permissão no [elemento ExtendedPermissions](/javascript/api/manifest/extendedpermissions) . Para obter detalhes, [consulte Implementar append-on-send em seu suplemento do Outlook](../outlook/append-on-send.md). Com o manifesto do Teams (versão prévia), você especifica essa permissão com o nome **Mailbox.AppendOnSend.User** em um objeto adicional na matriz "authorization.permissions.resourceSpecific".
> - Há uma permissão complementar necessária para suplementos que usam pastas compartilhadas. Com o manifesto XML, você especifica a permissão definindo [o elemento SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) como `true`. Para obter detalhes, consulte [Habilitar pastas compartilhadas e cenários de caixa de correio compartilhada em um suplemento do Outlook](../outlook/delegate-access.md). Com o manifesto do Teams (versão prévia), você especifica essa permissão com o nome **Mailbox.SharedFolder** em um objeto adicional na matriz "authorization.permissions.resourceSpecific".
