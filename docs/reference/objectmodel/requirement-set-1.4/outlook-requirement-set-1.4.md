# <a name="outlook-add-in-api-requirement-set-14"></a>Conjunto de requisitos de API para suplementos do Outlook versão 1.4

O subconjunto da API para suplementos do Outlook da API JavaScript para Office para inclui objetos, métodos, propriedades e eventos que você pode usar em um suplemento do Office.

> [!NOTE]
> Esta documentação destina-se a um [conjunto de requisitos](/javascript/office/requirement-sets/outlook-api-requirement-sets) diferente do conjunto de requisitos mais recente.

## <a name="whats-new-in-14"></a>Novidades na versão 1.4?

O conjunto de requisitos versão 1.4 inclui todos os recursos do [Conjunto de requisitos versão 1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md). O acesso ao namespace `Office.ui` foi adicionado.

### <a name="change-log"></a>Log de alterações

- [Office.context.ui.displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) adicionado: exibe uma caixa de diálogo em um host do Office.
- [Office.context.ui.messageParent](/javascript/api/office/office.ui#messageparent-messageobject-) adicionado: fornece uma mensagem da caixa de diálogo à sua página pai/de abertura.
- Objeto [Dialog](/javascript/api/office/office.dialog) adicionado: o objeto retornado quando o método  [`displayDialogAsync`](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) é chamado.

## <a name="see-also"></a>Confira também

- [Suplementos do Outlook](https://docs.microsoft.com/outlook/add-ins/)
- [Exemplos de código de suplementos do Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Introdução](https://docs.microsoft.com/outlook/add-ins/quick-start)