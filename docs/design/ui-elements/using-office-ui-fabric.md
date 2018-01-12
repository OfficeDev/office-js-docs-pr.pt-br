
#<a name="use-office-ui-fabric-261-in-office-add-ins"></a>Usar o Office UI Fabric 2.6.1 em suplementos do Office

Se estiver criando um suplemento do Office, recomendamos que você use o [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) para criar a experiência do usuário. As etapas a seguir fornecem orientação para as noções básicas de utilização do Fabric.  

> **Observação:** Para saber mais sobre o Office UI Fabric JS, confira [Usar o Office UI Fabric nos Suplementos do Office](https://dev.office.com/docs/add-ins/design/using-office-ui-fabric-js).

##<a name="1-set-up-fabric"></a>1. Configurar o Fabric
Adicione as seguintes linhas ao HTML na seção de cabeçalho para fazer referência ao Fabric na CDN.

     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.min.css">
     <link rel="stylesheet" href="https://appsforoffice.microsoft.com/fabric/1.0/fabric.components.min.css">


##<a name="2-use-fabric-icons-and-fonts"></a>2. Usar ícones e fontes do Fabric
É fácil usar ícones. Basta usar um elemento "i" e fazer referência às classes adequadas. Para controlar o tamanho do ícone, você pode alterar o tamanho da fonte.

    <i class="ms-Icon ms-Icon--group" style="font-size:xx-large" aria-hidden="true"></i>


##<a name="3-use-styles-for-simple-components"></a>3. Usar estilos para componentes simples
O Fabric vem com estilos para diversos elementos de interface do usuário, como botões e caixas de seleção. Basta fazer referência às classes adequadas para adicionar o estilo correspondente, conforme mostrado no exemplo a seguir.

    <button class="ms-Button" id="get-data-from-selection">
    <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
    <span class="ms-Button-label">Get Data from selection</span>
    <span class="ms-Button-description">Get Data from the document selection</span>
    </button>

##<a name="4-use-components-with-sample-behavior"></a>4. Usar componentes com modelo de comportamento
O Fabric inclui alguns componentes com suporte para comportamentos; por exemplo, o que ocorre no clique. Para começar, o **Fabric 2.6.1** inclui alguns **códigos de exemplo** que você pode usar, no formato de plug-ins de interface do usuário JQuery. Você pode também usar outras estruturas às quais pretende conectar itens. Se você optar por usar os exemplos, observe que o código não é distribuído como parte da CDN, portanto, você deve baixá-lo na versão **2.6.1** do [projeto do Fabric no GitHub](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1), referência a ele e, em seguida, inicializá-lo no código. 

Por exemplo, para usar o componente SearchBox:

1. Baixe o componente SearchBox do [GitHub](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1/src/components/SearchBox).
2. Adicione a seguinte referência ao código: `<script src="SearchBox/Jquery.SearchBox.js"></script>`
3. Para inicializar o componente, verifique se esta linha é executada quando a página for carregada: `$(".ms-SearchBox").SearchBox();`. Recomendamos incluir isso no bloco `Office.Initialize` do suplemento.     

**Observação:** Caso não pretenda usar todos os componentes do Fabric, você pode reduzir o tamanho dos recursos que baixar, optando por hospedar os arquivos CSS individuais de cada componente. Você pode obter os arquivos CSS das pastas de componentes no [repositório GitHub do Fabric 2.6.1](https://github.com/OfficeDev/office-ui-fabric-core/tree/release/2.6.1). 


##<a name="next-steps"></a>Próximas etapas
Se estiver procurando amostras de ponta a ponta que mostram como usar o Fabric, abordamos esse conteúdo para você. Confira a [Amostra de suplemento do Office do Fabric UI](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample). Se preferir, confira o site interativo do [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric).

