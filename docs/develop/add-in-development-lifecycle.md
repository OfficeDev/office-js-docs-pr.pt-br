
# <a name="office-add-ins-development-lifecycle"></a>Ciclo de vida de desenvolvimento de suplementos do Office

>
  **Observação:** Caso pretenda [publicar](../publish/publish.md) o suplemento na Office Store depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação da Office Store](https://msdn.microsoft.com/en-us/library/jj220035.aspx). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://msdn.microsoft.com/en-us/library/jj220035.aspx#Anchor_3) e a [Página de hospedagem e disponibilidade do suplemento do Office](https://dev.office.com/add-in-availability)).

O ciclo de vida de desenvolvimento típico de um Suplemento do Office inclui as seguintes etapas:


1.  **Decida qual é a proposta do suplemento.**
    
    Faça as seguintes perguntas:
    
      - Para quê o suplemento é útil? 
    
      - Como ele ajuda seus clientes a serem mais produtivos?
    
      - Quais cenários são compatíveis com os recursos do seu suplemento?
    

    Decida os recursos e cenários mais importantes, e concentre seu design nisso. 
    
2.  **Identifique os dados e a fonte de dados do suplemento.**
    
    Os dados estão em um documento, uma pasta de trabalho, uma apresentação, um projeto ou em um banco de dados do Access baseado no navegador? Há um item, ou itens, em um servidor Exchange ou caixa de correio do Exchange Online? Os dados vêm de uma fonte externa, como um serviço da Web?
    
3.  **Identifique o tipo de suplemento e os aplicativos host do Office que melhor dão suporte à finalidade do suplemento.**
    
    Considere o seguinte para identificar os cenários:
    
    - Os clientes usarão o suplemento para enriquecer o conteúdo de um documento ou um banco de dados baseado em navegador do Access? Em caso afirmativo, convém considerar a criação de um suplemento de conteúdo. 
    
    - Os clientes utilizarão o suplemento ao exibir ou ao escrever uma mensagem de email ou um compromisso? É importante poder expor o suplemento de acordo com o contexto atual? É uma prioridade disponibilizar o suplemento não apenas em computadores de mesa, mas também em tablets e telefones?
    
        Se a resposta for “Sim” para qualquer uma dessas perguntas, considere a criação de um suplemento do Outlook. Em seguida, identifique o contexto que acionará seu suplemento (por exemplo, o usuário está usando um formulário de composição, tipos de mensagem específicos, a presença de um anexo, um endereço, uma sugestão de tarefa ou de reunião, ou certos padrões de cadeia de caracteres no conteúdo de um compromisso ou um email). Confira [Regras de ativação para suplementos do Outlook](../outlook/manifests/activation-rules.md) para descobrir como é possível ativar o suplemento Outlook contextualmente.
    
    - Os clientes usarão o suplemento para aprimorar a experiência de criação ou de exibição de um documento? Em caso afirmativo, convém considerar a criação de um suplemento de painel de tarefas. 

    O suporte para determinadas APIs de suplementos pode ser diferente entre aplicativos do Office e de acordo com a plataforma em que estão sendo executados (no Windows, em Macs, na Web ou em dispositivos móveis). Para ver a cobertura da API atual pelo cliente e a plataforma, consulte nossa página [Disponibilidade de plataforma e host para o Suplemento do Office](https://dev.office.com/add-in-availability).  
    
4.  **Desenvolva e implemente a experiência do usuário e a interface do usuário para o suplemento.**
    
    Projete uma experiência de usuário rápida e fluida, que seja consistente, fácil de usar e com cenários primários que requerem apenas algumas etapas para serem executados. Dependendo da finalidade do suplemento, use APIs ou serviços da Web de terceiros.
    
    Você pode escolher entre várias ferramentas de desenvolvimento na Web e usar o HTML e JavaScript para implementar a interface do usuário.
    
5.  **Crie um arquivo de manifesto XML com base no esquema do manifesto dos Suplementos do Office.**
    
    Crie um manifesto XML para identificar o suplemento e seus requisitos, especificar os locais do HTML e de arquivos JavaScript e CSS que o suplemento possa vir a usar e, dependendo do tipo de suplemento, o tamanho e as permissões padrão.
    
    Para suplementos do Outlook, é possível especificar o contexto (com base na mensagem ou no compromisso atual) relevante para seu suplemento e que, portanto, faria o Outlook disponibilizá-lo na interface de usuário. Também é possível decidir quais dispositivos serão compatíveis com o suplemento. No manifesto, especifique o contexto para regras de ativação e dispositivos compatíveis.
    
6.  **Instale e teste o suplemento.**
    
    Coloque os arquivos HTML e todos os arquivos JavaScript e CSS nos servidores Web especificados no arquivo de manifesto do suplemento. O processo de instalação de um suplemento depende do tipo de suplemento. Para obter detalhes, confira [Realizar Sideload de Suplementos do Office para teste](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
    
    Para suplementos do Outlook, instale-os em uma caixa de correio do Exchange e especifique o local do arquivo de manifesto do suplemento no Centro de Administração do Exchange (EAC). Para saber mais, consulte [Implementar e instalar suplementos do Outlook para teste](../outlook/testing-and-tips.md).
    
7.  **Publique o suplemento.**
    
    Você pode enviar o suplemento para a Office Store, de onde os clientes podem instalá-lo. Além disso, publique os suplementos de painel de tarefas e de conteúdo em um catálogo de suplementos em uma pasta privada no SharePoint ou em uma pasta compartilhada na rede. Assim é possível implantar um suplemento do Outlook diretamente em um servidor do Exchange de sua organização. Para obter mais detalhes, veja [Publicar seu Suplemento do Office](../publish/publish.md).
    
8.  **Manter o suplemento**
    
    Se seu suplemento chama um serviço da Web e você fizer atualizações ao serviço da 
Web após a publicação do suplemento, não é preciso publicá-lo novamente. Entretanto, se você alterar quaisquer itens ou dados enviados para o suplemento, como o manifesto, capturas de tela, ícones, arquivos HTML ou JavaScript, é necessário publicar novamente o suplemento. Especificamente, se você publicar o suplemento na Office Store, será preciso enviá-lo novamente para que a Office Store possa implementar essas alterações. Você deve reenviar o suplemento com um manifesto atualizado que contenha um novo número de versão. Também deve garantir que o número de versão do suplemento seja atualizado no formulário de envio, de forma a corresponder ao novo número de versão do manifesto. Para suplementos do Outlook, certifique-se de que o elemento [Id](../../reference/manifest/id.md) contém uma UUID diferente no manifesto.
    
