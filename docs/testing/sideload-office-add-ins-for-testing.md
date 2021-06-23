---
title: Realizar sideload de suplementos do Office no Office na Web para teste
description: Teste seu Office de Office na Web ao fazer sideload.
ms.date: 04/14/2021
localization_priority: Normal
ms.openlocfilehash: e830ccbb6a4e325d6d70c3612492009b5e3d1570
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: pt-BR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077215"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>Realizar sideload de suplementos do Office no Office na Web para teste

Ao fazer sideload de um add-in, você pode instalar o add-in sem primeiro colocá-lo no catálogo de complementos. Isso é útil ao testar e desenvolver seu complemento porque você pode ver como o seu complemento aparecerá e funcionará.

Quando você faz sideload de um complemento, o manifesto do complemento é armazenado no armazenamento local do navegador, portanto, se você limpar o cache do navegador ou alternar para um navegador diferente, será preciso fazer o sideload do complemento novamente.

O sideload varia entre aplicativos host (por exemplo, Excel).

> [!NOTE]
> O sideload conforme descrito neste artigo tem suporte no Excel, OneNote, PowerPoint e Word. Para realizar o sideload de um suplemento do Outlook, confira [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Realizar sideload de um suplemento do Office no Office na Web

Esse processo é suportado apenas para **Excel,** **OneNote,** **PowerPoint** e **Word.** Para outros aplicativos host, consulte as instruções de sideload manual na seção a seguir. Este projeto de exemplo pressupo que você está usando um projeto criado com o gerador [Yeoman para Office Desempois](https://github.com/OfficeDev/generator-office).

1. Abra [Office na Web](https://office.live.com/). Usando a **opção Criar,** crie um documento em **Excel,** **OneNote,** **PowerPoint** ou **Word**. Neste novo documento, selecione **Compartilhar** na faixa de opções, selecione **Copiar Link** e copie a URL.

2. No diretório raiz dos arquivos do projeto yo do office, abra o **arquivopackage.json.** Na seção **config** deste arquivo, crie uma `"document"` propriedade. Colar a URL copiada como o valor da `"document"` propriedade. Por exemplo, o seu terá uma aparência assim:

    ```json
      "config": {
        "document": "<YOUR URL>",
        ...
      }
    ```

    > [!TIP]
    > Se você estiver criando um complemento que não está usando nosso gerador Yeoman, poderá adicionar parâmetros de consulta à URL do documento, acrescentando o seguinte à URL existente:

    - A porta do servidor de dev, como `&wdaddindevserverport=3000` .
    - O nome do arquivo de manifesto, como `&wdaddinmanifestfile=manifest1.xml` .
    - O GUID do manifesto, como `&wdaddinmanifestguid=05c2e1c9-3e1d-406e-9a91-e9ac64854143` .

    > Se você estiver usando o gerador Yeoman, adicionar essas informações não será necessário, pois a ferramenta Yeoman acrescenta essas informações automaticamente.
    > Observe que, em ambos os casos, no entanto, você só pode carregar manifestos de localhost.

3. Na linha de comando que começa no diretório raiz do seu projeto, execute o seguinte comando: `npm run start:web` .

4. Na primeira vez que você usar esse método para fazer sideload de um complemento na Web, você verá uma caixa de diálogo solicitando que você habilita o modo de desenvolvedor. Selecione a caixa de seleção Para **Habilitar o Modo de Desenvolvedor agora** e selecione **OK**.

5. Você verá uma segunda caixa de diálogo, perguntando se deseja registrar um manifesto de Office de complemento do seu computador. Você deve selecionar **Sim**.

6. Seu complemento está instalado. Se for um comando de complemento, ele deverá aparecer na faixa de opções ou no menu de contexto. Se for um complemento do painel de tarefas, o painel de tarefas deverá aparecer.

## <a name="sideload-an-office-add-in-in-office-on-the-web-manually"></a>Fazer sideload de Office de um Office na Web manualmente

Esse método não usa a linha de comando e pode ser realizado usando comandos somente no aplicativo host (como Excel).

1. Abra [Office na Web](https://office.live.com/). Abra um documento em **Excel,** **Word** ou **PowerPoint**. Na guia **Inserir** na faixa de opções na seção **Add-ins,** escolha **Office Adicionar.**

1. Na caixa **de diálogo Office de** Office, selecione a guia MEUS **ADD-INS,** escolha Gerenciar Meus **Complementos** e, em seguida, **Upload Meu Complemento**.

    ![A caixa Office de Office com um drop-down na leitura superior direita "Gerenciar meus complementos" e um drop-down abaixo dele com a opção "Upload Meu Complemento".](../images/office-add-ins-my-account.png)

1. **Navegue** até o arquivo de manifesto do suplemento e selecione **Carregar**.

    ![A caixa de diálogo Carregar suplemento com botões para pesquisar, carregar e cancelar.](../images/upload-add-in.png)

1. Verifique se o suplemento está instalado. Por exemplo, se for um comando do suplemento, ele deve aparecer na faixa de opções ou no menu de contexto. Se for um suplemento de painel de tarefas, o painel deve ser exibido.

> [!NOTE]
> Para testar seu Office de Microsoft Edge com o WebView (EdgeHTML) original, uma etapa de configuração adicional é necessária. Em um Windows de comando, execute a seguinte linha: `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` . Isso não é necessário quando o Office está usando o Chromium WebView2 baseado em Borda. Para obter mais informações, consulte [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="sideload-an-office-add-in"></a>Fazer sideload Office add-in

1. Entre na sua conta Microsoft 365 de usuário.

2. Abra o Iniciador de Aplicativos na extremidade esquerda da barra de ferramentas e selecione **Excel**, **Word** ou **PowerPoint** e crie um novo documento.

3. As etapas 3 a 6 são as mesmas da seção anterior **Realize sideload para um suplemento do Office no Office na Web**. 

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Sideload de um suplemento usando o Visual Studio

Se você estiver usando Visual Studio para desenvolver seu complemento, o processo de sideload será semelhante ao sideload manual para a Web. A única diferença é que você deve atualizar o valor do elemento **SourceURL** no manifesto para incluir a URL completa em que o suplemento for implantado.

> [!NOTE]
> Embora você possa realizar o sideload de suplementos do Visual Studio para o Office na Web, não é possível depurá-los no Visual Studio. Para depurar você precisará usar as ferramentas de depuração do navegador. Para saber mais, confira [Depurar suplementos no Office na Web](debug-add-ins-in-office-online.md).

1. No Visual Studio, abra a janela **Propriedades** escolhendo **Modo de exibição** > **Janela de propriedades**.
2. No **Gerenciador de Soluções**, selecione o projeto Web. Isso exibirá as propriedades para o projeto na janela **Propriedades**.
3. Na janela Propriedades, copie a **URL de SSL**.
4. No projeto de suplemento, abra o arquivo XML do manifesto. Certifique-se de que você está editando o XML do código-fonte. Para alguns tipos de projeto o Visual Studio abrirá o modo de exibição de visualização do XML que não funcionará para a próxima etapa.
5. Pesquisar e substituir todas as instâncias de **~remoteAppUrl/** pela URL de SSL que você copiou. Você verá várias substituições dependendo do tipo de projeto e as novas URLs serão muito similares a `https://localhost:44300/Home.html`.
6. Salve o arquivo XML.
7. Clique com botão direito do mouse no projeto Web e escolha **Depurar** > **Iniciar nova instância**. Isso executará o projeto Web sem iniciar o Office.
8. No Office na Web, realize o sideload do suplemento usando as etapas descritas anteriormente em [Sideload de um suplemento do Office no Office na Web](#sideload-an-office-add-in-in-office-on-the-web).

## <a name="remove-a-sideloaded-add-in"></a>Remover um complemento com sideload

Você pode remover um complemento com sideload anteriormente limpando o cache do navegador. Se você fizer alterações no manifesto do seu complemento (por exemplo, atualizar nomes de arquivos de ícones ou texto de comandos de complemento), talvez seja necessário limpar o [cache do Office](clear-cache.md) e, em seguida, re-sideload do complemento usando o manifesto atualizado. Isso permitirá que o Office processe o suplemento conforme descrito no manifesto atualizado.

## <a name="see-also"></a>Confira também

- [Fazer sideload de Suplementos do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md)
- [Realizar sideload de suplementos do Outlook para teste](../outlook/sideload-outlook-add-ins-for-testing.md)
- [Limpar o cache do Office](clear-cache.md)
