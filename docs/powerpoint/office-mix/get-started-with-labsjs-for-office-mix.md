
# <a name="get-started-with-labsjs-for-office-mix"></a>Introdução ao LabsJS para o Office Mix



O conteúdo do LabsJS expõe uma API (labs.js), amostras, documentação e arquivos associados que você pode usar para desenvolver laboratórios interativos, integrá-los ao Office Mix e depois processá-los no Microsoft PowerPoint. Na verdade, esses laboratórios são Suplementos do Office que você cria usando HTML5 e a biblioteca JavaScript labs.js.

## <a name="labsjs-content"></a>Conteúdo do LabsJS

O LabsJS fornece a documentação, exemplos de laboratório e os arquivos necessários para criar e publicar seus próprios laboratórios para o Office Mix.


**Arquivos necessários**


|**File**|**Descrição**|
|:-----|:-----|
|labs-1.0.4.js|A API JavaScript LabsJS para o desenvolvimento de laboratórios do Office Mix. Este arquivo deve ser incluído em seu projeto para permitir a integração com o Office Mix. O arquivo também é hospedado em uma CDN (rede de distribuição de conteúdo) em <code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>. Quando você publica seu aplicativo, é necessário criar um vínculo com o arquivo na CDN.|
|labs-1.0.4.d.ts|Arquivo de definição em TypeScript para labs.js. Isso possibilita uma integração fácil com seu código TypeScript com o labs.js. O arquivo de definição também fornece uma visão geral de todos os componentes contidos em labs.js. Você pode baixar o TypeScript em [http://www.typescriptlang.org/](http://www.typescriptlang.org/). O arquivo de definição foi compilado com base no TypeScript versão 0.9.1.1.|
|Histórico|Histórico de versões das várias versões da biblioteca labs.js.|
|Labshost.html|Uma página da Web que permite a exibição e depuração de seu laboratório no Office Mix, fora do contexto do PowerPoint. Para usar a página, digite a URL para a caixa de entrada principal e ela será carregada dentro do quadro. Dados trocados entre a API e o Office Mix durante a execução no PowerPoint ou no reprodutor de lição do Office Mix aparecerão nas caixas de inserção à direita. Os dados também podem ser previamente propagados. Observe que os exemplos de Laboratórios na seção Exemplos mostram os Suplementos do Office Mix existentes em execução no contexto do host.|
|SampleManifest.xml|Uma amostra de manifesto dos Suplementos do Office para usar como um modelo para a criação de seu próprio manifesto de aplicativo.|
|Simplelab.html|Um exemplo de Laboratório criado com labs.js. Permite a seleção e a inserção de uma página da Web, que depois rastreará o usuário que a visualiza.|
|Simplelab.ts|O arquivo TypeScript usado para criar o exemplo de simplelab.|
|Simplelab.js|Versão JavaScript do exemplo de Simplelab. Esse arquivo e o simplelab.ts mostram o uso da API LabsJS.|

## <a name="set-up-your-development-environment"></a>Definir seu ambiente de desenvolvimento

A biblioteca labs.js serve como uma camada de abstração sobre a biblioteca office.js (a API para Suplementos do Office), para que os laboratórios que você criar usando a biblioteca labs.js sejam realmente Suplementos do Office. Para trabalhar com a biblioteca labs.js e executar esses laboratórios dentro do Office Mix, primeiro defina-se como um desenvolvedor de Suplementos do Office.


### <a name="register-for-an-office-365-developer-site"></a>Registrar-se em um Site do Desenvolvedor do Office 365

A primeira etapa é inscrever-se em um Site do Desenvolvedor do Office 365. Isso permite que você hospede e teste seu laboratório antes de enviá-lo à Office Store. O site permite que você publique seu suplemento no Office Mix e teste-o em um ambiente ativo.

Para saber mais, confira [Configurar um ambiente de desenvolvimento para Suplementos do SharePoint no Office 365](http://msdn.microsoft.com/library/b22ce52a-ae9e-4831-9b68-c9210af6dc54%28Office.15%29.aspx). 


### <a name="set-up-an-app-catalog-on-sharepoint-online"></a>Configurar um catálogo de aplicativos no SharePoint Online

Após a criação e provisionamento de seu site de desenvolvedor, configure um catálogo de suplementos no SharePoint Online. Para saber mais, confira [Configurar um catálogo de suplementos no Office 365](../../publish/publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md).

Para o Office Mix, use um catálogo de suplementos para que você possa inserir suplementos de pré-produção em uma lição e conduzir testes completos antes de enviar os laboratórios para o repositório.


## <a name="create-your-lab"></a>Criar seu laboratório

Para criar seu primeiro laboratório, execute as etapas no [Passo a passo](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md), que explica como criar um questionário simples com opções de verdadeiro/falso. Confira [Passo a passo: criar o seu primeiro laboratório para o Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)


## <a name="publish-your-lab"></a>Publicar seu laboratório

Depois de criar seu laboratório, você pode publicá-lo e enviá-lo para o repositório.


### <a name="create-and-upload-your-application-manifest"></a>Criar e carregar o manifesto do aplicativo

O manifesto de aplicativo é um documento XML que descreve seu laboratório LabJS. Ele fornece uma referência para a URL na qual o laboratório está hospedado e fornece detalhes sobre o laboratório, incluindo o nome de exibição, a descrição, os ícones, o tamanho etc.

Incluímos um exemplo de manifesto, "SampleManifest.xml". Para saber mais sobre o esquema de manifesto e obter um link para a definição do esquema, confira [Manifesto XML dos suplementos do Office](../../overview/add-in-manifests.md).

Para carregar seu manifesto em seu site do SharePoint, primeiro vá para seu catálogo de aplicativos, que geralmente está na URL <code>https://\<your site\>/sites/AppCatalog</code>. Depois, escolha o botão **Novo aplicativo** e siga as etapas para carregar seu manifesto de aplicativo.


### <a name="update-your-powerpoint-2013-catalog"></a>Atualizar seu catálogo do PowerPoint 2013

Em seguida, atualize seu catálogo do PowerPoint 2013 Depois disso, você pode entrar com sua conta de desenvolvedor.

Comece atualizando seu catálogo do PowerPoint 2013. Inicie o PowerPoint 2013 e navegue pelo caminho de menu **Arquivo > Opções > Central de Confiabilidade > Configurações da Central de Confiabilidade > Catálogos de Aplicativos Confiáveis**. A partir daí, adicione uma referência ao seu catálogo de aplicativos e escolha **OK**. O PowerPoint 2013 pedirá que você saia para que as mudanças tenham efeito. Saia.

Por fim, entre novamente usando a conta de desenvolvedor. Escolha seu nome de logon no canto superior direito no PowerPoint 2013 e faça logon usando sua conta de desenvolvedor. Agora você pode inserir seu suplemento.


### <a name="insert-publish-and-view-your-app"></a>Inserir, publicar e exibir seu aplicativo

Para inserir seu suplemento no catálogo, escolha a faixa de opções **Inserir** e depois escolha **Store** na seção **Aplicativos**. Escolha **Minha Organização** e você verá o suplemento em seu catálogo de suplementos. Escolha o suplemento, selecione **Inserir** e seu suplemento (laboratório) será inserido no documento do PowerPoint 2013.

Agora você pode tirar proveito de todas as funcionalidades disponíveis do Office Mix para publicar sua lição com seu novo laboratório.


 >**Importante**:  Para exibir o aplicativo, você precisa fazer logon em seu catálogo do SharePoint com o mesmo navegador usado para exibir suas lição. Os catálogos do SharePoint apenas permitem o acesso de usuários autenticados. Portanto, para ver seu aplicativo você precisa fazer logon primeiro. 


### <a name="submit-your-lab-to-the-office-store"></a>Enviar seu laboratório à Office Store

Para enviar seu laboratório à Office Store, confira [Publicar seu Suplemento do Office](../../publish/publish.md)


## <a name="additional-resources"></a>Recursos adicionais



- [Suplementos do Office Mix](../../powerpoint/office-mix/office-mix-add-ins.md)
    
- [Suplementos do Office](../../overview/office-add-ins.md)
    
- [Criando o seu primeiro laboratório para o Office Mix](../../powerpoint/office-mix/creating-your-first-lab-for-office-mix.md)
    
