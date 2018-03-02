---
title: Usar o Office UI Fabric JS em Suplementos do Office
description: ''
ms.date: 12/04/2017
---

# <a name="use-office-ui-fabric-js-in-office-add-ins"></a>Usar o Office UI Fabric JS em Suplementos do Office

O Office UI Fabric é uma estrutura de front-end JavaScript destinada à criação de experiências de usuário para Office e Office 365. Se você criar um suplemento usando somente JavaScript, sem usar uma estrutura como Angular ou React, considere o uso do Fabric JS para criar a experiência do usuário. Para saber mais, confira [Office UI Fabric JS](https://dev.office.com/fabric-js).

Este artigo explica as noções básicas do uso do Fabric JS.  

## <a name="add-the-fabric-cdn-references"></a>Adicionar as referências de CDN do Fabric
Para fazer referência ao Fabric a partir da CDN, adicione o seguinte código HTML à página.

```html
<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
<link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css">
<script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>
```

## <a name="use-fabric-js-ux-components"></a>Usar os componentes da experiência de usuário do Fabric JS

O Fabric JS fornece diversos componentes da experiência de usuário, como botões e caixas de seleção, que você pode usar no suplemento. Veja a seguir uma lista de componentes da experiência de usuário do Fabric JS recomendados para uso em suplementos. Para usar um dos componentes do Fabric no suplemento, siga o link para a documentação do Fabric e siga as instruções do tópico **Usar este componente**. 

- [Navegação estrutural](https://dev.office.com/fabric-js/Components/Breadcrumb/Breadcrumb.html)
- [Botão](https://dev.office.com/fabric-js/Components/Button/Button.html) (considere o uso da variante de botão pequeno no suplemento. Adicione 16px de preenchimento para pequenos botões para garantir um destino de toque mínimo de 40px em dispositivos de toque.)
- [Caixa de seleção](https://dev.office.com/fabric-js/Components/CheckBox/CheckBox.html)
- [ChoiceFieldGroup](https://dev.office.com/fabric-js/Components/ChoiceFieldGroup/ChoiceFieldGroup.html)
- [Seletor de data](https://dev.office.com/fabric-js/Components/DatePicker/DatePicker.html) (Para um exemplo que mostra como implementar o seletor de data em um suplemento, confira o exemplo de código do [Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).)
- [Lista suspensa](https://dev.office.com/fabric-js/Components/Dropdown/Dropdown.html)
- [Rótulo](https://dev.office.com/fabric-js/Components/Label/Label.html)
- [Link](https://dev.office.com/fabric-js/Components/Link/Link.html)
- [Lista](https://dev.office.com/fabric-js/Components/List/List.html) (Considere alterar os estilos padrão do componente no CSS.)
- [MessageBanner](https://dev.office.com/fabric-js/Components/MessageBanner/MessageBanner.html)
- [MessageBar](https://dev.office.com/fabric-js/Components/MessageBar/MessageBar.html)
- [Sobreposição](https://dev.office.com/fabric-js/Components/Overlay/Overlay.html)
- [Painel](https://dev.office.com/fabric-js/Components/Panel/Panel.html)
- [Pivô](https://dev.office.com/fabric-js/Components/Pivot/Pivot.html)
- [ProgressIndicator](https://dev.office.com/fabric-js/Components/ProgressIndicator/ProgressIndicator.html)
- [Caixa de Pesquisa](https://dev.office.com/fabric-js/Components/SearchBox/SearchBox.html)
- [Controle giratório](https://dev.office.com/fabric-js/Components/Spinner/Spinner.html)
- [Tabela](https://dev.office.com/fabric-js/Components/Table/Table.html)
- [TextField](https://dev.office.com/fabric-js/Components/TextField/TextField.html)
- [Alternância](https://dev.office.com/fabric-js/Components/Toggle/Toggle.html)
   
## <a name="updating-your-add-in-to-use-fabric-js"></a>Atualizar o suplemento para usar o Fabric JS
Se você já está usando uma versão anterior do Office UI Fabric e pretende migrar para o Fabric JS, familiarize-se, incorpore e teste os novos componentes no suplemento. Para ajudá-lo a planejar as atualizações, lembre-se do seguinte:

- A inicialização de componentes é mais simples usando o Fabric JS. Em relação às versões anteriores do Fabric, você deve incluir o arquivo JavaScript do componente do Fabric no projeto do suplemento, incluído em uma referência de `<Script>` a esse arquivo, e inicializar o componente. No Fabric JS, não é mais necessário incluir o arquivo JavaScript do componente do Fabric e a referência de `<Script>` associada. Tudo o que você precisa fazer é inicializar o componente do Fabric.   
- Agora, vários componentes fornecem funções que controlam o comportamento do componente da experiência de usuário. Por exemplo, o controle da caixa de seleção tem uma função `toggle` que permite alternar entre os estados marcado e desmarcado. 
- Atualizamos alguns estilos e nomes de classe de ícones.
- A alteração mais significativa é o uso do elemento `<label>` em vários componentes. O elemento `<label>` controla o estilo do componente. Pode ser necessário atualizar o código da experiência de usuário para usar o elemento `<label>`. Por exemplo, alterar o valor do atributo verificado `<input>` do elemento em uma caixa de seleção do Fabric JS não afeta a caixa de seleção. Em vez disso, use as funções `check`, `unCheck` ou `toggle`.   

## <a name="implementation"></a>Implementação
Se estiver procurando um exemplo de código de ponta a ponta que mostra como usar o Fabric JS, abordamos esse conteúdo para você. Confira o seguinte recurso:

- [Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

## <a name="see-also"></a>Veja também
Se estiver procurando exemplos de códigos ou documentação de uma versão anterior do Fabric, confira as seguintes opções:

- [Padrões de design da experiência de usuário (usa o Fabric 2.6.1)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Amostra de Fabric UI do suplemento do Office (usa o Fabric 1.0)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Usar o Fabric 2.6.1 em um Suplemento do Office](ui-elements/using-office-ui-fabric.md)
 

