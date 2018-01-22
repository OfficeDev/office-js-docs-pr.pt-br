
# <a name="use-the-oauth-authorization-framework-in-an-office-add-in"></a>Usar a estrutura de autorização OAuth em um Suplemento do Office

OAuth é o padrão aberto para autorização que provedores de serviço online como Office 365, Facebook, Google, SalesForce, LinkedIn e outros usam para executar a autenticação do usuário. A estrutura de autorização OAuth é o protocolo de autorização padrão usado no Azure e no Office 365. A estrutura de autorização OAuth é usada em cenários empresariais (corporativos) e de consumidor.

Os provedores de serviços online podem fornecer APIs públicas expostas via REST. Os desenvolvedores podem usar essas APIs públicas em suplementos do Office para ler ou gravar dados para o provedor de serviços online. A integração de dados de provedores de serviços online em um suplemento aumenta seu valor, o que leva a uma maior adoção pelos usuários. Ao usar essas APIs em seu suplemento, os usuários deverão fazer a autenticação usando a estrutura de autorização OAuth.

Este tópico descreve como implementar um fluxo de autenticação no suplemento para executar a autenticação do usuário. Os segmentos de código incluídos neste tópico são obtidos do exemplo de código [Office-Add-in-NodeJS-ServerAuth](https://github.com/OfficeDev/Office-Add-in-NodeJS-ServerAuth).

 **Observação** Por motivos de segurança, os navegadores não têm permissão para exibir páginas de entrada em um IFrame. Dependendo da versão do Office que seus clientes usam, principalmente versões baseadas na Web, o suplemento é exibido em um IFrame. Isso impõe algumas considerações sobre como gerenciar o fluxo de autenticação. 

O diagrama a seguir mostra os componentes necessários e o fluxo de eventos que ocorrem durante a implementação da autenticação no suplemento.

![Realizar uma autenticação OAuth em um Suplemento do Office](../images/OAuthInOfficeAddin.png)

O diagrama mostra como os seguintes componentes necessários são usados:


- O Office executa um suplemento de painel de tarefas no computador do usuário. O suplemento abre uma janela pop-up para iniciar o fluxo de autenticação. Os suplementos não podem iniciar fluxos de autenticação diretamente porque, dependendo da plataforma usada, os suplementos podem ser executados em um IFRAME. Por motivos de segurança, páginas de entrada OAuth não podem ser exibidas em um IFRAME. 
    
- Um servidor Web hospeda o código do suplemento. Este exemplo de código usa um servidor de banco de dados em execução no servidor Web para armazenar o token de acesso do usuário. É necessário persistir o token de acesso para que, depois que a autenticação for concluída usando a janela pop-up, as páginas do suplemento principal possam usar os mesmos tokens para acessar dados do serviço online. É necessário salvar os tokens usando opções no servidor porque você não pode depender de informações passadas do suplemento ou do pop-up.
    
- O provedor OAuth 2.0 executa a autenticação do usuário.
    

    
 **Importante** Os tokens de acesso não podem ser retornados ao painel de tarefas, mas podem ser usados no servidor. Neste exemplo de código, os tokens de acesso são armazenados no banco de dados por dois minutos. Após dois minutos, os tokens são limpos do banco de dados e os usuários são solicitados a realizar a autenticação novamente. Antes de alterar esse período de tempo em sua própria implementação, considere os riscos de segurança associados ao armazenamento de tokens de acesso em um banco de dados por um período de tempo de mais de dois minutos.


## <a name="step-1---start-socket-and-open-a-pop-up-window"></a>Etapa 1 ‒ iniciar o soquete e abrir uma janela pop-up

Quando você executa este código de exemplo, um suplemento de painel de tarefas é exibido no Office. Quando o usuário escolhe um provedor OAuth no qual fazer logon, primeiro o suplemento cria um soquete. Este exemplo usa um soquete para fornecer uma boa experiência do usuário no suplemento. O suplemento usa o soquete para comunicar o sucesso ou a falha da autenticação ao usuário. Com o uso de um soquete, a página principal do suplemento é facilmente atualizada com o status de autenticação e não requer interação com o usuário nem sondagem. O segmento de código a seguir, obtido de routes/connect.js, mostra como iniciar o soquete. O soquete é nomeado usando **decodedNodeCookie**, que é a ID de sessão do suplemento Este exemplo de código cria o soquete usando [socket.io](http://socket.io/).


```js
io.on('connection', function (socket) {
  console.log('Socket connection established');
  var jsonCookie =
    cookie.parse(socket
      .handshake
      .headers
      .cookie);
  var decodedNodeCookie =
    cookieParser
      .signedCookie(jsonCookie.nodecookie, '<Insert a random string>');
  console.log('Decoded cookie: ' + decodedNodeCookie);
  // The session ID becomes the room name for this session.
  socket.join(decodedNodeCookie);
  io.to(decodedNodeCookie).emit('init', 'Private socket session established');
});

```

Em seguida, o suplemento se conecta ao soquete. O código a seguir pode ser encontrado em /public/javascripts/client.js.




```js
var socket = io.connect('https://localhost:3001', { secure: true });
```

Em seguida, o suplemento abre uma janela pop-up no computador do usuário usando **window.open**. Ao executar **window.open**, verifique se o URI de redirecionamento e a ID de sessão do suplemento são passados na URL. A ID de sessão do suplemento é usada para identificar o soquete a ser usado ao enviar informações de status de autenticação à interface do usuário do suplemento. O segmento de código a seguir pode ser encontrado em views/index.jade.




```js
onclick="window.open('/connect/azure/#{sessionID}', 'AuthPopup', 'width=500,height=500,centerscreen=1,menubar=0,toolbar=0,location=0,personalbar=0,status=0,titlebar=0,dialog=1')")
```


## <a name="steps-2-amp-3---start-the-authentication-flow-and-show-the-sign-in-page"></a>Etapas 2 &amp; 3 ‒ iniciar o fluxo de autenticação e mostrar a página de entrada

O suplemento deve iniciar o fluxo de autenticação. O segmento de código abaixo usa a biblioteca Passport OAuth. Ao iniciar o fluxo de autenticação, passe a URL de autorização do provedor OAuth e a ID de sessão do suplemento. A ID de sessão do suplemento deve ser passada no parâmetro de estado. Agora a janela pop-up exibe a página de entrada do provedor OAuth para que os usuários possam entrar.


```js
router.get('/azure/:sessionID', function(req, res, next) { 
   passport.authenticate( 
     'azure',  
     { state: req.params.sessionID }, 

```


## <a name="steps-4-5-amp-6---user-signs-in-and-web-server-receives-tokens"></a>Etapas 4, 5 &amp; 6 ‒ o usuário entra e o servidor Web recebe tokens

 Após uma entrada bem-sucedida, um token de acesso, um token de atualização e um parâmetro de estado são retornados para o suplemento. O parâmetro de estado contém a ID de sessão, que é usada para enviar informações de status de autenticação ao soquete na etapa 7. O segmento de código a seguir, obtido de app.js, armazena o token de acesso no banco de dados.


```js
  dbHelperInstance.insertDoc(userData, null, 
         function (err, body) { 
           if (!err) { 
             console.log("Inserted session entry [" + userData.sessid + "] id: " + body.id); 
           } 
           done(err, userData); 
         }); 

```


## <a name="step-7---show-authentication-information-in-the-add-ins-ui"></a>Etapa 7 ‒ mostrar informações de autenticação na interface do usuário do suplemento

O segmento de código a seguir, obtido de connect.js, atualiza interface do usuário do suplemento com as informações de status de autenticação. A interface do usuário do suplemento é atualizada usando o soquete que foi criado na etapa 1.


```js
  
       io.to(user.sessid).emit('auth_success', providers); 
       next(); 

```


## <a name="additional-resources"></a>Recursos adicionais
<a name="bk_addresources"> </a>


- [Exemplo de Autenticação do Servidor de Suplemento do Office para Node.js](https://github.com/OfficeDev/Office-Add-in-Nodejs-ServerAuth/blob/master/README.md)
    
