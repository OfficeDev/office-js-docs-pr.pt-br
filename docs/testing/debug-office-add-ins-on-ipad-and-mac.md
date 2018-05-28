---
title: Depurar suplementos do Office no iPad e no Mac
description: ''
ms.date: 03/21/2018
ms.openlocfilehash: 5d68fa000e19d81ebbcd1b383a790958f2bbac72
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: pt-BR
ms.lasthandoff: 05/23/2018
---
# <a name="debug-office-add-ins-on-ipad-and-mac"></a>Depurar suplementos do Office no iPad e no Mac

Voc? pode usar o Visual Studio para desenvolver e depurar suplementos no Windows, mas n?o pode us?-lo para depurar suplementos no iPad ou no Mac. Como os suplementos s?o desenvolvidos usando HTML e Javascript, s?o projetados para funcionar em v?rias plataformas, mas pode haver diferen?as sutis em como cada navegador processa o HTML. Este artigo descreve como depurar suplementos em execu??o em um iPad ou em um Mac. 

## <a name="debugging-with-safari-web-inspector-on-a-mac"></a>Depura??o com o Safari Web Inspector em um Mac

Voc? pode depurar um suplemento do Office usando o Safari Web Inspector. 

Para poder depurar suplementos do Office no Mac, voc? deve ter o Mac OS High Sierra e Mac Office Vers?o: 16.9.1 (compila??o 18012504) ou posterior. Se voc? n?o tiver uma compila??o do Office Mac, poder? obter uma ao adquirir o [programa Office 365 Developer](https://aka.ms/o365devprogram).

Para come?ar, abra um terminal e defina a propriedade `OfficeWebAddinDeveloperExtras` para o aplicativo relevante do Office da seguinte maneira:

- `defaults write com.microsoft.Word OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Powerpoint OfficeWebAddinDeveloperExtras -bool true`

- `defaults write com.microsoft.Outlook OfficeWebAddinDeveloperExtras -bool true`

Em seguida, abra o aplicativo do Office e insira seu suplemento. Clique com o bot?o direito no suplemento e voc? ver? a op??o **Inspecionar elemento** no menu de contexto.  Selecione essa op??o e ela abrir? o Inspetor, onde voc? pode definir pontos de interrup??o e depurar seu suplemento.

> [!NOTE]
> Observe que esse ? um recurso experimental e n?o h? garantias de que preservaremos essa funcionalidade em vers?es futuras de aplicativos do Office.

## <a name="debugging-with-vorlonjs-on-a-ipad-or-mac"></a>Depura??o com o Vorlon.JS em um iPad ou Mac

Para depurar um suplemento no iPad ou Mac, voc? pode usar o Vorlon.JS, um depurador para p?ginas da Web que ? semelhante ?s ferramentas F12. Ele ? projetado para funcionar remotamente e permite depurar p?ginas da Web em dispositivos diferentes. Para saber mais, veja o [site do Vorlon](http://www.vorlonjs.com).  


### <a name="install-and-set-up-vorlonjs"></a>Instalar e configurar o Vorlon.JS  

1.  Fa?a logon no dispositivo como um administrador.

2.  Instale o [Node.js](https://nodejs.org) se ele ainda n?o estiver instalado. 

3.  Abra uma janela do **Terminal** e digite o comando `npm i -g vorlon`. A ferramenta est? instalada em `/usr/local/lib/node_modules/vorlon`.


### <a name="configure-vorlonjs-to-use-https"></a>Configurar o Vorlon.JS para usar HTTPS

Para depurar um aplicativo usando o Vorlon.JS, adicione uma marca `<script>` ? p?gina de abertura do aplicativo que carrega um script Vorlon.JS de um local conhecido (veja os detalhes no procedimento a seguir). Se um suplementos for protegido por SSL (HTTPS), todos os scripts usados dever?o estar hospedados em um servidor HTTPS, inclusive o script Vorlon.JS. Portanto, voc? precisar? configurar o Vorlon.JS para usar SSL se quiser usar esse script com suplementos. 

> [!IMPORTANT]
> [!include[HTTPS guidance](../includes/https-guidance.md)]

1.  No **Localizador**, acesse `/usr/local/lib/node_modules/vorlon`, abra o menu de contexto (clique com o bot?o direito do mouse) da pasta `/Server` e escolha **Obter Informa??es**.

2.  Escolha o ?cone de cadeado no canto inferior direito da janela **Informa??es do servidor** para desbloquear a pasta.

3. Na se??o **Compartilhamento e Permiss?es** da janela, defina o **Privil?gio** para o grupo **funcion?rios** como **Leitura/Grava??o**.

4. Escolha o ?cone de cadeado novamente para ***voltar a bloquear*** a pasta.

5. No **Localizador**, expanda a subpasta `/Server`, clique com bot?o direito no arquivo `config.json` e selecione **Obter Informa??es**.

6. Na janela **informa??es de config.json**, altere os privil?gios do arquivo da mesma forma que voc? fez para sua pasta `/Server` pai. N?o se esque?a de bloquear novamente e de fechar a janela.

7. No **Localizador**, clique com bot?o direito do mouse no arquivo `config.json`, selecione **Abrir com**e selecione **TextEdit**. O arquivo ? aberto em um editor de texto.

8. Altere a propriedade **useSSL** para `true`.

9. Na se??o **plug-ins**, localize o plug-in com a **id** de `OFFICE` e o **nome** de `Office Addin`. Se a propriedade **enabled** do plug-in ainda n?o estiver como `true`, defina-a como `true`.

10. Salve o arquivo e feche o editor.

11. No **Localizador**, navegue at? `/usr/local/lib/node_modules/vorlon`, clique com bot?o direito do mouse na subpasta `Server` e selecione **Novo terminal na pasta**. 
    
12. Na janela do **Terminal**, digite `sudo vorlon`. Ser? solicitado que voc? digite sua senha de administrador. O servidor Vorlon ? iniciado. Deixe aberta a janela do **Terminal**.

13. Abra uma janela do navegador e v? para `https://localhost:1337`, que ? a interface do Vorlon.JS. Quando solicitado, escolha **Sempre** para confiar no certificado de seguran?a. 

    > [!NOTE]
    > Se n?o for solicitado, talvez seja necess?rio confiar no certificado manualmente. O arquivo de certificado ? `/usr/local/lib/node_modules/vorlon/Server/cert/server.crt`. Experimente as etapas a seguir. Se voc? tiver problemas, veja a ajuda do Macintosh ou do iPad. 
    >
    > 1. Feche a janela do navegador e na janela do **Terminal** que est? executando o servidor Vorlon, use Control-C para parar o servidor.
    > 2. No **Localizador**, clique com bot?o direito do mouse no arquivo `server.crt` e escolha **Acesso ao Conjunto de Chaves**. A janela **Acesso ao Conjunto de Chaves** ? exibida.
    > 3. Na lista **Conjuntos de Chaves** ? esquerda, escolha **logon**, caso ainda n?o estiver marcado, e, em seguida, escolha **Certificados** na se??o **Categoria**. Verifique se o **localhost** do certificado est? na lista.
    > 4. Clique com bot?o direito do mouse no **localhost** do certificado e escolha **Obter Informa??es**. Uma janela do **localhost** ? exibida.
    > 5. Na se??o **Confiar**, abra o seletor rotulado como **Ao usar este certificado** e escolha **Sempre Confiar**. 
    > 6. Feche a janela do **localhost**. Se a a??o for bem-sucedida, o certificado do **localhost** na janela **Acesso ao Conjunto de Chaves** exibir? uma cruz branca em um c?rculo azul no ?cone.


### <a name="configure-the-add-in-for-vorlonjs-debugging"></a>Configurar o suplemento para depura??o do Vorlon.JS

1. Adicione a seguinte marca de script ? se??o `<head>` do arquivo home.html (ou arquivo HTML principal) do seu suplemento:

    ```html
    <script src="https://localhost:1337/vorlon.js"></script>    
    ```  

2. Implante o aplicativo da Web do suplemento em um servidor Web que pode ser acessado do Mac ou iPad, como um site do Azure. 

3. Atualize a URL do suplemento em todos os locais onde a URL aparece no manifesto do suplemento.

4. No Mac ou iPad, copie o manifesto do suplemento na seguinte pasta: `/Users/{your_name_on_the_device}/Library/Containers/com.microsoft.{host_name}/Data/Documents/wef`, onde *{nome_do_host}* ? Word, Excel, PowerPoint ou Outlook.


### <a name="inspect-an-add-in-in-vorlonjs"></a>Inspecionar um suplemento no Vorlon.JS

1. Se o servidor Vorlon n?o estiver sendo executado, no **Localizador**, navegue at? `/usr/local/lib/node_modules/vorlon`, clique com bot?o direito na subpasta `Server` e selecione **Novo terminal na pasta**. 
    
2.  Na janela do **Terminal**, digite `sudo vorlon`. Ser? solicitado que voc? digite sua senha de administrador. O servidor Vorlon ? iniciado. Deixe aberta a janela do **Terminal**.

3.  Abra uma janela do navegador e v? para `https://localhost:1337`, que ? a interface do Vorlon.JS.

4. Realize o sideload do suplemento. Para o Excel, PowerPoint ou Word, realize o sideload conforme descrito em [Realizar sideload de um suplemento do Office no iPad e no Mac](sideload-an-office-add-in-on-ipad-and-mac.md). Se for um suplemento do Outlook, realize o sideload conforme descrito em [Realizar sideload de suplementos do Outlook para teste](https://docs.microsoft.com/en-us/outlook/add-ins/sideload-outlook-add-ins-for-testing). Se o suplemento n?o usar comandos de suplemento, ele ser? imediatamente aberto. Caso contr?rio, escolha o bot?o para abrir o suplemento. Dependendo da compila??o do aplicativo host do Office, o bot?o ser? exibido em ambas guias **P?gina Inicial** ou em uma guia **Suplemento**.

O suplemento aparecer? na lista de Clientes no Vorlon.JS (no lado esquerdo da interface do Vorlon.JS) como **{OS} - n**, para um determinado n?mero *n* e onde *{OS}* ? o tipo de dispositivo, como "Macintosh". 

![Captura de tela que mostra a interface do Vorlon.js](../images/vorlon-interface.png)

A ferramenta Vorlon tem uma variedade de plug-ins. Os que estiverem habilitados no momento ser?o exibidos como guias na parte superior da ferramenta. (? poss?vel habilitar mais plug-ins escolhendo o ?cone de engrenagem no canto esquerdo). Esses plug-ins s?o semelhantes ?s fun??es nas ferramentas F12. Por exemplo, voc? pode real?ar elementos DOM, executar comandos e muito mais. Para obter mais detalhes, veja [Principais plug-ins da documenta??o do Vorlon](http://vorlonjs.com/documentation/#console) 

Um plug-in do **Suplemento do Office** adiciona recursos extras ao Office.js, como explorar o modelo de objeto e executar chamadas de Office.js e ler os valores das propriedades de objetos. Para obter instru??es, veja [Plug-in do VorlonJS para depura??o de suplementos do Office](https://blogs.msdn.microsoft.com/mim/2016/02/18/vorlonjs-plugin-for-debugging-office-addin/).

> [!NOTE]
> N?o ? poss?vel definir pontos de interrup??o no Vorlon.JS.


## <a name="clearing-the-office-applications-cache-on-a-mac-or-ipad"></a>Limpar cache do aplicativo do Office em um Mac ou iPad

Os Suplementos muitas vezes s?o armazenados em cache no Office para Mac por quest?o de desempenho. Normalmente, o cache ser? limpo quando o suplemento for recarregado. Se houver mais de um suplemento no mesmo documento, ? prov?vel que o processo de limpeza autom?tica do cache ao recarregar n?o seja confi?vel. 

No Mac, o cache pode ser limpo manualmente ao excluir tudo na pasta `/Users/{your_name_on_the_device}/Library/Containers/com.Microsoft.OsfWebHost/Data/`. 

No iPad, voc? pode chamar `window.location.reload(true)` a partir do JavaScript no suplemento para for?ar uma recarrega. Uma outra alternativa ? reinstalar o Office.
