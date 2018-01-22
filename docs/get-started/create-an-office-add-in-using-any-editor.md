
# <a name="create-an-office-add-in-using-any-editor"></a>Criar um Suplemento do Office usando qualquer editor

Você pode usar o gerador Yeoman para criar seu Suplemento do Office. O gerador Yeoman oferece o gerenciamento de compilação e estrutura do projeto. O arquivo `manifest.xml` informa ao aplicativo do Office onde o suplemento está localizado e como você deseja que ele seja mostrado. O aplicativo do Office se encarrega de hospedá-lo no Office.

 >**Observação:** essas instruções usam o Terminal em um Mac, mas você também pode usar outros ambientes de shell. 


## <a name="prerequisites-for-the-yeoman-generator"></a>Pré-requisitos para o gerador Yeoman

Para instalar o gerador Yeoman do Office, instale o [git](https://git-scm.com/downloads) e o node.js no seu computador. Se estiver em um Mac, recomendamos que você use o [Gerenciador de Versão de Nós](https://github.com/creationix/nvm) para instalar o node.js com as permissões corretas. Se estiver no Windows, instale o node.js de [nodejs.org](https://nodejs.org/en/).

>**Observação:** se estiver no Windows, use os valores padrão ao instalar seu git, com as seguintes exceções:

>- Usar o git no prompt de comando do Windows
>- Usar a janela de console padrão do Windows

Depois de instalar o node.js, abra um Terminal e instale o gerador globalmente.

```
npm install -g yo generator-office
```


## <a name="create-the-default-files-for-your-add-in"></a>Criar os arquivos padrão para o suplemento

O gerador Yeoman é executado no diretório em que você deseja manter a estrutura do projeto. Antes de desenvolver um Suplemento do Office, primeiro você deve criar uma pasta para o projeto.

No Terminal, vá para a pasta pai em que deseja criar o projeto. Em seguida, use estes comandos para criar uma nova pasta chamada _myHelloWorldaddin_ e mudar o diretório atual para ela:




```
mkdir myHelloWorldaddin
cd myHelloWorldaddin
```

Use o gerador Yeoman para criar o suplemento de sua escolha. As etapas neste artigo criam um suplemento simples do painel tarefas. Para executar o gerador, insira o seguinte comando:




```
yo office
```

**Entrada do gerador Yeoman para um suplemento**

O gerador solicitará o seguinte: 


- Nova subpasta – use _N_
- Nome do suplemento – use _myHelloWorldaddin_ 
- O aplicativo do Office compatível – escolha qualquer aplicativo
- Crie um novo suplemento – use _Sim, quero um novo suplemento._
- Adicione [TypeScript](https://www.typescriptlang.org/) – use _N_
- Escolha uma estrutura – use _Jquery_

>**Observação:** se você quiser criar um Suplemento do Office que use o Office UI Fabric React, insira o seguinte:
>- Adicione [TypeScript](https://www.typescriptlang.org/) – use _Y_
>- Escolha uma estrutura – use _React_

![Gif do gerador Yeoman solicitando uma entrada do projeto](../images/gettingstarted-fast.gif)

Isso cria a estrutura e os arquivos básicos para o suplemento.


## <a name="hosting-your-office-add-in"></a>Hospedar o suplemento do Office

Os suplementos do Office devem ser hospedados, mesmo em desenvolvimento, via HTTPS. O Yo Office cria um bsconfig.json, que usa o Browsersync para que você possa ajustar e testar mais rapidamente seu suplemento sincronizando as alterações nos arquivos em vários dispositivos. 

Inicie o site HTTPS local em https://localhost:3000 digitando o seguinte comando no seu console:


```
npm start
```

O Browsersync iniciará um servidor HTTPS e inicializará o arquivo index.html no seu projeto. Você verá uma mensagem de erro informando "Há um problema com o certificado de segurança deste site".


![Gif mostrando o processo para ignorar o erro e ver o arquivo de index.html padrão](../images/ssl-chrome-bypass.gif)

Esse erro ocorre porque o Browsersync contém um certificado SSL autoassinado no qual seu ambiente de desenvolvimento deve confiar. Confira mais informações sobre como resolver esse erro em [Adicionar certificados autoassinados](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

## <a name="sideload-the-add-in-into-office"></a>Realizar o sideload do suplemento no Office

Você pode usar sideloading para instalar o suplemento para teste nos clientes do Office:

- [Realizar o sideload de suplementos do Office para teste](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
- [Realizar o sideload de suplementos do Office em um iPad ou Mac para teste](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)   
- [Realizar o sideload de suplementos do Outlook para teste](../outlook/testing-and-tips.md)

## <a name="develop-your-office-add-in"></a>Desenvolver o suplemento do Office

Use qualquer editor de texto para desenvolver os arquivos para o suplemento do Office personalizado.

> **Importante:** o arquivo manifest-myHelloWorldaddin.xml informa aos aplicativos clientes do Office como interagir com seu suplemento. O valor na marca `<id>` é um GUID que o Yo Office cria quando gera o projeto. Não altere o GUID do suplemento. Se o host for o Microsoft Azure, o valor de `SourceLocation` será uma URL semelhante a _https://[nome-do-aplicativo-Web].azurewebsites.net/[caminho-para-o-suplemento]_. Se estiver usando a opção auto-hospedado, como neste exemplo, ela será _https://localhost:3000/[caminho-para-o-suplemento]_.


## <a name="debug-your-office-add-in"></a>Depurar o suplemento do Office


É possível depurar o suplemento de várias maneiras:

- Anexe um depurador do painel de tarefas (Office 2016 para Windows).
- Use as ferramentas de desenvolvedor do seu navegador.
- Use as ferramentas de desenvolvedor F12 no Windows 10.

### <a name="attach-debugger-from-the-task-pane"></a>Anexe o depurador do painel de tarefas

No Office 2016 para Windows, Build 77xx.xxxx ou posterior, é possível anexar o depurador do painel de tarefas. O recurso de anexar o depurador anexará diretamente o depurador ao processo correto do Internet Explorer. É possível anexar um depurador independentemente de você estar utilizando Yeoman Generator, Visual Studio Code, node.js, Angular ou outra ferramenta. 

Para saber mais, confira [Anexar depurador do painel de tarefas](../testing/attach-debugger-from-task-pane.md).


### <a name="browser-developer-tools"></a>Pesquisar ferramentas de desenvolvedor 

Você pode usar clientes Web do Office e abrir as ferramentas de desenvolvedor do navegador para depurar o seu suplemento como faz com qualquer outro aplicativo JavaScript no lado do cliente. 

### <a name="f12-developer-tools-on-windows-10"></a>Ferramentas de desenvolvedor F12 no Windows 10

Se estiver usando o cliente para área de trabalho do Office no Windows 10, você pode [Depurar suplementos usando as ferramentas de desenvolvedor F12 no Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).
    
## <a name="next-steps"></a>Próximas etapas

- [Implantar e publicar seu suplemento do Office](../publish/publish.md)
    
