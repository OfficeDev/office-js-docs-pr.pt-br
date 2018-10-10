# <a name="outlook-javascript-api-requirement-sets"></a>Conjuntos de requisitos da API JavaScript do Outlook

Os suplementos do Outlook declaram quais versões de API eles exigem usando o elemento [Requirements](/javascript/office/manifest/requirements) em seu [manifesto](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests). Os suplementos do Outlook sempre incluem um elemento [Set](/javascript/office/manifest/set) com um atributo `Name` definido como `Mailbox` e um atributo `MinVersion` definido como o conjunto de requisitos mínimos de API compatível com os cenários do suplemento.

Por exemplo, o seguinte trecho do manifesto indica um conjunto de requisitos mínimos de 1.1:

```xml
<Requirements>
  <Sets>
    <Set Name="MailBox" MinVersion="1.1" />
  </Sets>
</Requirements>
```

Todas as APIs do Outlook pertencem ao `Mailbox` [conjunto de requisitos](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements). O conjunto de requisitos `Mailbox` tem versões e cada novo conjunto de APIs que lançamos pertence a uma versão superior. Nem todos os clientes do Outlook serão compatíveis com o conjunto mais recente de APIs, mas se um cliente do Outlook declarar suporte a um conjunto de requisitos, ele será compatível com todas as APIs desse conjunto.

A especificação de uma versão mínima de conjunto de requisitos no manifesto controla em quais clientes do Outlook o suplemento aparecerá. Se um cliente não for compatível com o conjunto de requisitos mínimos, ele não carregará o suplemento. Por exemplo, se for especificada a versão 1.3 do conjunto de requisitos, o suplemento não aparecerá nos clientes do Outlook incompatíveis com a versão 1.3.

## <a name="using-apis-from-later-requirement-sets"></a>Usar APIs de conjuntos de requisitos posteriores

A especificação de um conjunto de requisitos não limita as APIs disponíveis que podem ser usadas pelo suplemento. Por exemplo, se o suplemento especificar o conjunto de requisitos 1.1, mas for executado em um cliente do Outlook compatível com a versão 1.3, o suplemento poderá usar as APIs do conjunto de requisitos 1.3.

Para usar as APIs mais recentes, os desenvolvedores podem apenas verificar sua existência usando a técnica JavaScript padrão:

```js
if (item.somePropertyOrFunction !== undefined) {
  item.somePropertyOrFunction ...
}
```

Essas verificações não são necessárias para APIs que estão presentes na versão do conjunto de requisitos especificada no manifesto.

## <a name="choosing-a-minimum-requirement-set"></a>Escolher um conjunto de requisitos mínimos

Os desenvolvedores devem usar o conjunto de requisitos mínimos que contém o conjunto essencial de APIs para seu cenário, sem o qual o suplemento não funcionará.

## <a name="clients"></a>Clientes

Os clientes a seguir são compatíveis com os suplementos do Outlook.

| Cliente | Conjuntos de requisitos de API compatíveis |
| --- | --- |
| Outlook 2019 para Windows | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2019 para Mac | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2016 (Click-to-Run) para Windows | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6), [1.7](/javascript/office/objectmodel/requirement-set-1.7/outlook-requirement-set-1.7) |
| Outlook 2016 (MSI) para Windows | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook 2016 para Mac | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook 2013 para Windows | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4) |
| Outlook para iPhone | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook para Android | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5) |
| Outlook na Web (Office 365 e Outlook.com) | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3), [1.4](/javascript/office/objectmodel/requirement-set-1.4/outlook-requirement-set-1.4), [1.5](/javascript/office/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5), [1.6](/javascript/office/objectmodel/requirement-set-1.6/outlook-requirement-set-1.6) |
| Outlook Web App (Exchange 2013 no local) | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1) |
| Outlook Web App (2016 do Exchange no local) | [1.1](/javascript/office/objectmodel/requirement-set-1.1/outlook-requirement-set-1.1), [1.2](/javascript/office/objectmodel/requirement-set-1.2/outlook-requirement-set-1.2), [1.3](/javascript/office/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3) |

> [!NOTE]
> O suporte à versão 1.3 no Outlook 2013 foi adicionado como parte da [atualização do Outlook 2013 de 8 de dezembro de 2015 (KB3114349)](https://support.microsoft.com/kb/3114349). O suporte à versão 1.4 no Outlook 2013 foi adicionado como parte da [atualização do Outlook 2013 de 13 de dezembro de 2016 (KB3118280)](https://support.microsoft.com/help/3118280).
