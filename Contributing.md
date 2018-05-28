# <a name="contribute-to-this-documentation"></a>Contribuir para esta documenta??o

Agradecemos seu interesse em nossa documenta??o!

* [Maneiras de contribuir](#ways-to-contribute)
* [Contribuir usando o GitHub](#contribute-using-github)
* [Contribuir usando o Git](#contribute-using-git)
* [Como usar o Markdown para formatar seu t?pico](#how-to-use-markdown-to-format-your-topic)
* [Perguntas frequentes](#faq)
* [Mais recursos](#more-resources)

## <a name="ways-to-contribute"></a>Maneiras de contribuir

Veja a seguir algumas maneiras de contribuir com esta documenta??o:

* Para fazer pequenas altera??es em um artigo [contribua usando o GitHub](#contribute-using-github).
* Para fazer grandes altera??es ou altera??es que envolvam c?digos, [contribua usando o Git](#contribute-using-git).
* Relatar bugs na documenta??o por meio do GitHub Issues
* Solicitar nova documenta??o no site [UserVoice de Plataforma do Desenvolvedor do Office](http://officespdev.uservoice.com)

## <a name="contribute-using-github"></a>Contribuir usando o GitHub

Use o GitHub para contribuir com esta documenta??o sem precisar clonar o reposit?rio em sua ?rea de trabalho. Essa ? a maneira mais f?cil de criar uma solicita??o pull neste reposit?rio. Use este m?todo para fazer uma pequena altera??o que n?o envolva altera??es de c?digo. 

**Observa??o** Usar este m?todo permite contribuir em um artigo de cada vez.

### <a name="to-contribute-using-github"></a>Para contribuir usando o GitHub

1. Localize o artigo com o qual deseja contribuir no GitHub. 

    Se o artigo estiver no MSDN, escolha o link **sugerir e enviar altera??es**, na se??o **Contribuir com este conte?do**, e voc? ser? direcionado ao mesmo artigo no GitHub.
2. Quando estiver no artigo no GitHub, acesse o GitHub (obtenha uma conta gratuita em [Junte-se ao GitHub](https://github.com/join)).
3. Escolha o **?cone de l?pis** (editar o arquivo em sua bifurca??o deste projeto) e fa?a suas altera??es na janela **<> Edit fie**. 
4. Role at? a parte inferior e insira a descri??o.
5. Escolha a op??o para propor a altera??o e criar a solicita??o pull em **Propose file change**>**Create pull request**.

Assim, voc? envia com ?xito uma solicita??o pull. As solicita??es pull geralmente s?o analisadas dentro de 10 dias ?teis. 


## <a name="contribute-using-git"></a>Contribuir usando o Git

Use o Git para fazer altera??es substanciais, tais como:

* Contribuir com c?digos.
* Contribuir com altera??es que afetam o significado.
* Contribuir com grandes altera??es de texto.
* Adicionar novos t?picos.

### <a name="to-contribute-using-git"></a>Para contribuir usando o Git

1. Se voc? n?o tiver uma conta, configure uma no [GitHub](https://github.com/join). 
2. Depois que tiver a conta, instale o Git em seu computador. Siga os passos em [Configurando o Tutorial do Git](https://help.github.com/articles/set-up-git/).
3. Para enviar uma solicita??o pull usando o Git, siga as etapas da se??o [Usar o GitHub, o Git e este reposit?rio](#use-github-git-and-this-repository).
4. Ser? solicitado que voc? assine o Contrato de licen?a do colaborador se voc? for:

    * um membro do grupo Microsoft Open Technologies;
    * colaboradores que n?o trabalham na Microsoft.

Como membro da comunidade, voc? deve assinar o Contrato de Licen?a de Contribui??o (CLA) antes de poder contribuir com envios volumosos para um projeto. Voc? s? precisa completar e enviar a documenta??o uma vez. Reveja cuidadosamente o documento. Talvez seja necess?rio que seu empregador assine o documento.

A assinatura do Contrato de Licen?a de Contribui??o (CLA) n?o lhe concede direito a confirmar altera??es no reposit?rio principal, mas isso significa que as equipes do Office Developer e do Office Developer Content Publishing poder?o revisar e aprovar suas contribui??es. Voc? ser? creditado por seus envios.

As solicita??es pull geralmente s?o analisadas dentro de 10 dias ?teis.

## <a name="use-github-git-and-this-repository"></a>Use o GitHub, o Git e este reposit?rio

**Observa??o:** A maior parte das informa??es desta se??o pode ser encontrada nos artigos de [Ajuda do GitHub].  Se voc? estiver familiarizado com o Git e o GitHub, pule para a se??o **Contribuir e editar conte?do** para ver as informa??es espec?ficas sobre o fluxo de c?digo/conte?do desse reposit?rio.

### <a name="to-set-up-your-fork-of-the-repository"></a>Configurar sua bifurca??o do reposit?rio

1.  Configure uma conta GitHub para que voc? pode contribuir para esse projeto. Caso ainda n?o tenha feito isso, acesse o [GitHub](https://github.com/join) e fa?a isso agora.
2.  Instale o Git em seu computador. Siga os passos de [Configurando o Tutorial do Git] [Set Up Git].
3.  Crie sua pr?pria bifurca??o desse reposit?rio. Para fazer isso, escolha o bot?o **Bifurca??o** localizado na parte superior da p?gina.
4.  Copie sua bifurca??o para seu computador. Para fazer isso, abra o Git Bash. No prompt de comando, digite:

        git clone https://github.com/<your user name>/<repo name>.git

    Em seguida, crie uma refer?ncia para o reposit?rio raiz inserindo esses comandos:

        cd <repo name>
        git remote add upstream https://github.com/OfficeDev/<repo name>.git
        git fetch upstream

Parab?ns! Agora seu reposit?rio est? configurado. Voc? n?o precisar? repetir essas etapas novamente.

### <a name="contribute-and-edit-content"></a>Contribuir e editar o conte?do

Para que o processo de contribui??o seja o mais cont?nuo poss?vel, siga estas etapas.

#### <a name="to-contribute-and-edit-content"></a>Para contribuir e editar conte?do

1. Crie uma nova ramifica??o.
2. Adicione novo conte?do ou edite o conte?do existente.
3. Envie uma solicita??o pull para o reposit?rio principal.
4. Exclua a ramifica??o.

**Importante**: limite cada ramifica??o a um ?nico conceito/artigo para simplificar o fluxo de trabalho e reduzir a chance de conflitos de mesclagem. O conte?do apropriado para uma nova ramifica??o inclui:

* um novo artigo;
* edi??es de ortografia e gram?tica; e
* aplicar uma ?nica altera??o de formata??o em um grande conjunto de artigos (por exemplo, aplicar um novo rodap? sobre direito autoral).

#### <a name="to-create-a-new-branch"></a>Para criar uma nova ramifica??o

1.  Abra o Git Bash.
2.  No prompt de comando do Git Bash, digite: `git pull upstream master:<new branch name>`. Isso cria uma nova ramifica??o local copiada da ?ltima ramifica??o-mestra do OfficeDev.
3.  No prompt de comando do Git Bash, digite: `git push origin <new branch name>`. Isso alertar? o GitHub para a nova ramifica??o. Agora voc? dever? surgir a nova ramifica??o na sua bifurca??o do reposit?rio no GitHub.
4.  No prompt de comando do Git Bash, digite `git checkout <new branch name>` para alternar para a nova ramifica??o.

#### <a name="add-new-content-or-edit-existing-content"></a>Adicionar novo conte?do ou editar o conte?do existente

Navegue at? o reposit?rio em seu computador usando o Explorador de Arquivos. Os arquivos do reposit?rio estar?o em `C:\Users\<yourusername>\<repo name>`.

Para editar arquivos, abra-os em um editor de sua escolha e modifique-os. Para criar um novo arquivo, use o editor de sua escolha e salve o novo arquivo no local apropriado em sua c?pia local do reposit?rio. Enquanto estiver trabalhando, salve seu trabalho com frequ?ncia.

Os arquivos localizados no `C:\Users\<yourusername>\<repo name>` s?o uma c?pia de trabalho da ramifica??o nova que voc? criou em seu reposit?rio local. Qualquer que seja a altera??o voc? fa?a nessa pasta, ela s? afetar? o reposit?rio local quando voc? confirmar uma altera??o. Para confirmar uma altera??o no reposit?rio local, digite os seguintes comandos no GitBash:

    git add .
    git commit -v -a -m "<Describe the changes made in this commit>"

O comando `add` adiciona suas altera??es para uma ?rea de preparo em prepara??o para confirm?-las no reposit?rio. O per?odo posterior ao comando `add` especifica que voc? deseja preparar todos os arquivos adicionados ou modificados, verificando repetidamente as subpastas. (Caso voc? n?o queira confirmar todas as altera??es, ? poss?vel adicionar arquivos espec?ficos. Voc? tamb?m pode desfazer uma confirma??o. Para obter ajuda, digite `git add -help` ou `git status`.)

O comando `commit` aplica as altera??es preparadas ao reposit?rio. A op??o `-m` significa que voc? est? fornecendo o coment?rio de confirma??o na linha de comando. As op??es -v e -a podem ser omitidas. A op??o -v corresponde ? sa?da detalhada do comando e a op??o -a faz o que voc? j? fez com o comando adicionar.

Voc? pode confirmar v?rias vezes enquanto estiver fazendo seu trabalho ou apenas uma vez quando terminar.

#### <a name="submit-a-pull-request-to-the-main-repository"></a>Enviar uma solicita??o pull para o reposit?rio principal.

Quando terminar o trabalho e estiver pronto para mescl?-lo no reposit?rio principal, siga estas etapas.

#### <a name="to-submit-a-pull-request-to-the-main-repository"></a>Para enviar uma solicita??o pull para o reposit?rio principal

1.  No prompt de comando do Git Bash, digite `git push origin <new branch name>`. Em seu reposit?rio local, `origin` refere-se ao reposit?rio do GitHub a partir do qual voc? clonou o reposit?rio local. Esse comando coloca o estado atual do sua nova ramifica??o, incluindo todas as confirma??es feitas nas etapas anteriores, na ramifica??o do GitHub.
2.  No site do GitHub, navegue em sua bifurca??o para a nova ramifica??o.
3.  Escolha o bot?o **Pull Request** na parte superior da p?gina.
4.  Verifique se o branch Base ? `OfficeDev/<repo name>@master` e o branch Head ? `<your username>/<repo name>@<branch name>`.
5.  Escolha o bot?o para atualiza o intervalo de confirma??o **Update Commit Range**.
6.  Inclua um t?tulo ? sua solicita??o pull e descreva todas as altera??es que voc? estiver fazendo.
7.  Envie a solicita??o pull.

Um dos administradores do site processar? sua solicita??o pull. Sua solicita??o pull ficar? vis?vel no site OfficeDev/<repo name> em Issues. Quando a solicita??o pull for aceita, o problema ser? resolvido.

#### <a name="create-a-new-branch-after-merge"></a>Criar uma nova ramifica??o ap?s a mesclagem

Depois que uma ramifica??o for mesclada com sucesso (ou seja, sua solicita??o for aceita), n?o continue a trabalhar na ramifica??o local. Isso poder? gerar conflitos de mesclagem caso voc? envie outra solicita??o pull. Para fazer uma nova atualiza??o, crie uma nova ramifica??o local com base na ramifica??o de upstream mesclada com ?xito e, ent?o, exclua a ramifica??o local inicial.

Por exemplo, se sua ramifica??o local X foi mesclada com ?xito na ramifica??o-mestra OfficeDev/microsoft-graph-docs e voc? quer fazer atualiza??es adicionais no conte?do mesclado. Crie uma nova ramifica??o local, X2, da ramifica??o-mestra OfficeDev/microsoft-graph-docs. Para fazer isso, abra o GitBash e execute os seguintes comandos:

    cd microsoft-graph-docs
    git pull upstream master:X2
    git push origin X2

Agora voc? tem c?pias locais (em uma nova ramifica??o local) do trabalho que enviou na ramifica??o X. A ramifica??o X2 tamb?m cont?m todo o trabalho que outros autores mesclaram, portanto, se seu trabalho depender do trabalho de outras pessoas (por exemplo, imagens compartilhadas), ele estar? dispon?vel em nova ramifica??o. Voc? pode confirmar se seu trabalho anterior (e o trabalho de outras pessoas) est? na ramifica??o verificando a nova ramifica??o...

    git checkout X2

... e verificando o conte?do. (O comando `checkout` atualiza os arquivos no `C:\Users\<yourusername>\microsoft-graph-docs` para o estado atual da ramifica??o do X2.) Assim que voc? verificar um novo branch, ser? poss?vel fazer atualiza??es no conte?do e confirm?-las como de costume. No entanto, para evitar trabalhar na ramifica??o mesclada (X) por engano, o melhor a fazer ser? exclu?-la (confira a seguinte se??o: **Excluir uma ramifica??o**).

#### <a name="delete-a-branch"></a>Excluir uma ramifica??o

Depois que as altera??es forem mescladas com ?xito no reposit?rio principal, exclua a ramifica??o utilizada, pois voc? n?o precisar? mais dela.  Qualquer trabalho adicional deve ser feito em uma nova ramifica??o.  

#### <a name="to-delete-a-branch"></a>Para excluir uma ramifica??o

1.  No prompt de comando do Git Bash, digite `git checkout master`. Isso garante que voc? n?o fique na ramifica??o a ser exclu?da (o que n?o ? permitido).
2.  Em seguida, no prompt de comando, digite `git branch -d <branch name>`. Isso exclui a ramifica??o em seu computador somente se ela tiver sido mesclada com ?xito no reposit?rio upstream. (Voc? pode substituir esse comportamento com o sinalizador `?D`, mas primeiro certifique-se de que deseja fazer isso).
3.  Por fim, digite `git push origin :<branch name>` no comando prompt (um espa?o antes dos dois pontos e nenhum espa?o depois deles).  Essa a??o excluir? a ramifica??o em sua bifurca??o no github.  

Parab?ns, voc? contribuiu com ?xito para o projeto.

## <a name="how-to-use-markdown-to-format-your-topic"></a>Como usar o Markdown para formatar seu t?pico

### <a name="standard-markdown"></a>Markdown-padr?o

Todos os artigos neste reposit?rio usam Markdown. A introdu??o completa (e a listagem de toda a sintaxe) pode ser encontrada na [P?gina Inicial do Markdown](http://daringfireball.net/projects/markdown/ 
).
 
## <a name="faq"></a>Perguntas frequentes

### <a name="how-do-i-get-a-github-account"></a>Como obter uma conta do GitHub?

Preencha o formul?rio em [Ingressar no GitHub](https://github.com/join) para abrir uma conta gratuita do GitHub. 

### <a name="where-do-i-get-a-contributors-license-agreement"></a>Onde posso obter um Contrato de Licen?a do Colaborador? 

Um aviso ser? automaticamente enviado para voc? informando que ? preciso assinar o CLA (Contrato de Licen?a do Colaborador) se sua solicita??o de recebimento exigir um. 

Como membro da comunidade, **voc? deve assinar o CLA (Contrato de Licen?a do Colaborador) antes de poder contribuir com envios volumosos para esse projeto**. Voc? s? precisa concluir e enviar a documenta??o uma vez. Reveja cuidadosamente o documento. Talvez seja necess?rio que seu empregador assine o documento.

### <a name="what-happens-with-my-contributions"></a>O que acontece com as minhas contribui??es?

Quando voc? envia suas altera??es, por meio de uma solicita??o pull, nossa equipe ser? notificada e a examinar?. Voc? receber? notifica??es sobre sua solicita??o pul do GitHub. Al?m disso, voc? tamb?m poder? ser notificado por uma pessoa de nossa equipe se precisarmos de mais informa??es. Se a solicita??o de recep??o for aprovada, atualizaremos a documenta??o. Reservamo-nos o direito de editar seu envio por motivos legais, estil?sticos, de clareza ou por outros problemas.

### <a name="can-i-become-an-approver-for-this-repositorys-github-pull-requests"></a>Posso me tornar um aprovador de solicita??es pull desse reposit?rio do GitHub?

Atualmente, n?o estamos autorizando que colaboradores externos aprovem solicita??es pull neste reposit?rio.

### <a name="how-soon-will-i-get-a-response-about-my-change-request"></a>Em quanto tempo terei uma resposta sobre a solicita??o de altera??o?

As solicita??es pull geralmente s?o analisadas dentro de 10 dias ?teis.


## <a name="more-resources"></a>Mais recursos

* Para saber mais sobre o Markdown, acesse o site do criador do Markdown [Daring Fireball].
* Para saber mais sobre como usar o Git e o GitHub, primeiro confira a [se??o de Ajuda do GitHub] [Ajuda do GitHub].

[GitHub Home]: http://github.com
[Ajuda do GitHub]: http://help.github.com/
[Set Up Git]: http://help.github.com/win-set-up-git/
[Markdown Home]: http://daringfireball.net/projects/markdown/
[Daring Fireball]: http://daringfireball.net/
