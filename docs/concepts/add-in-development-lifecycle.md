---
title: Ciclo de vida de desenvolvimento de suplementos do Office
description: ''
ms.date: 01/23/2018
---

# <a name="office-add-ins-development-lifecycle"></a>Ciclo de vida de desenvolvimento de suplementos do Office

> [!NOTE]
> Caso pretenda [publicar](../publish/publish.md) o suplemento na experiência do Office depois de criá-lo, verifique se você está em conformidade com as [Políticas de validação do AppSource](https://docs.microsoft.com/pt-br/office/dev/store/validation-policies). Por exemplo, para passar na validação, seu suplemento deve funcionar em todas as plataformas com suporte aos métodos que você definir (para mais informações, confira a [seção 4.12](https://docs.microsoft.com/pt-br/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) e a [Página de hospedagem e disponibilidade de suplementos do Office](../overview/office-add-in-availability.md)). 

O ciclo de vida de desenvolvimento típico de um Suplemento do Office inclui as seguintes etapas:


## <a name="1-decide-on-the-purpose-of-the-add-in"></a>1. Decida qual é a proposta do suplemento
    
Faça as seguintes perguntas:
    
- Para quê o suplemento é útil? 
        
- Como ele ajuda seus clientes a serem mais produtivos?
        
- Quais cenários são compatíveis com os recursos do seu suplemento?
    
Decida os recursos e cenários mais importantes e concentre seu design nisso. 

    
## <a name="2-identify-the-data-and-data-source-for-the-add-in"></a>2. Identifique os dados e a fonte de dados do suplemento
    
- Os dados em um documento, pasta de trabalho, apresentação, projeto ou um banco de dados do Access são baseados em navegador? 
    
- Os dados sobre um item ou itens estão no Exchange Server ou em uma caixa de correio do Exchange Online? 
    
- Os dados são provenientes de uma fonte externa, como um serviço Web?

    
## <a name="3-identify-the-type-of-add-in-and-office-host-applications-that-best-support-the-purpose-of-the-add-in"></a>3. Identifique o tipo de suplemento e os aplicativos host do Office que dão o melhor suporte à finalidade do suplemento.
    
Considere o seguinte para identificar os cenários:
    
- Os clientes usarão o suplemento para enriquecer o conteúdo de um documento ou um banco de dados baseado em navegador do Access? Em caso afirmativo, convém considerar a criação de um **suplemento de conteúdo**. 
    
- Os clientes utilizarão o suplemento ao exibir ou ao escrever uma mensagem de email ou um compromisso? É importante poder expor o suplemento de acordo com o contexto atual? É uma prioridade disponibilizar o suplemento não apenas em computadores de mesa, mas também em tablets e telefones?
    
    Se a resposta for Sim para qualquer uma dessas perguntas, considere a criação de um **suplemento do Outlook**. Identifique o contexto que acionará seu suplemento (por exemplo, o usuário está usando um formulário de composição, tipos de mensagem específicos, a presença de um anexo, um endereço, uma sugestão de tarefa ou de reunião, ou certos padrões de cadeia de caracteres no conteúdo de um compromisso ou um email). 
        
    Para descobrir como é possível ativar o suplemento Outlook contextualmente, confira as [Regras de ativação para suplementos do Outlook](https://docs.microsoft.com/pt-br/outlook/add-ins/activation-rules). 
    
- Os clientes usarão o suplemento para aprimorar a experiência de criação ou de exibição de um documento? Em caso afirmativo, convém considerar a criação de um **suplemento de painel de tarefas**. 

O suporte para determinadas APIs de suplementos pode ser diferente entre aplicativos do Office e de acordo com a plataforma em que estão sendo executados (no Windows, no Mac, na Web ou em dispositivos móveis). Para ver a cobertura da API atual pelo cliente e a plataforma, veja nossa página [Disponibilidade de plataforma e host para o Suplemento do Office](../overview/office-add-in-availability.md).  

    
## <a name="4-design-and-implement-the-user-experience-and-user-interface-for-the-add-in"></a>4. Desenvolva e implemente a experiência do usuário e a interface do usuário para o suplemento.
    
Projete uma experiência de usuário rápida e fluida, que seja consistente, fácil de usar e com cenários primários que requerem apenas algumas etapas para serem executados. Dependendo da finalidade do suplemento, use APIs ou serviços da Web de terceiros.
    
Você pode escolher entre várias ferramentas de desenvolvimento na Web e usar o HTML e JavaScript para implementar a interface do usuário.

    
## <a name="5-create-an-xml-manifest-file-based-on-the-office-add-ins-manifest-schema"></a>5. Crie um arquivo de manifesto XML com base no esquema do manifesto dos Suplementos do Office.
    
Crie um manifesto XML para identificar o suplemento e seus requisitos, especificar os locais do HTML e de arquivos JavaScript e CSS que o suplemento possa vir a usar e, dependendo do tipo de suplemento, o tamanho e as permissões padrão.
    
Para suplementos do Outlook, é possível especificar o contexto (com base na mensagem ou no compromisso atual) relevante para seu suplemento e que, portanto, faria o Outlook disponibilizá-lo na interface de usuário. Também é possível decidir quais dispositivos serão compatíveis com o suplemento. No manifesto, especifique o contexto para regras de ativação e dispositivos compatíveis.
    

## <a name="6-install-and-test-the-add-in"></a>6. Instale e teste o suplemento.
    
Coloque os arquivos HTML e todos os arquivos JavaScript e CSS nos servidores Web especificados no arquivo de manifesto do suplemento. O processo de instalação de um suplemento depende do tipo de suplemento. Para obter detalhes, confira [Realizar Sideload de Suplementos do Office para teste](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
    
Para suplementos do Outlook, instale-os em uma caixa de correio do Exchange e especifique o local do arquivo de manifesto do suplemento no Centro de Administração do Exchange (EAC). Para saber mais, consulte [Implementar e instalar suplementos do Outlook para teste](https://docs.microsoft.com/pt-br/outlook/add-ins/testing-and-tips).

    
## <a name="7-publish-the-add-in"></a>7. Publique o suplemento.
    
Você pode enviar o suplemento no AppSource, de onde os clientes podem instalá-lo. Além disso, publique os suplementos de painel de tarefas e de conteúdo em um catálogo de suplementos em uma pasta privada no SharePoint ou em uma pasta compartilhada na rede. Assim é possível implantar um suplemento do Outlook diretamente em um servidor do Exchange de sua organização. Para obter mais detalhes, veja [Publicar seu Suplemento do Office](../publish/publish.md).
    
    
## <a name="8-maintain-the-add-in"></a>8. Faça a manutenção do suplemento
    
Se o suplemento chamar um serviço Web e você atualizar o serviço Web depois de publicar o suplemento, não será preciso publicar o suplemento novamente. No entanto, se você alterar os itens ou dados enviados ao suplemento, por exemplo, o manifesto do suplemento, capturas de tela, ícones, arquivos HTML ou JavaScript, você precisará publicá-lo novamente. 
    
Especificamente, se você publicar o suplemento no AppSource, será preciso reenviar o suplemento para que o AppSource possa implementar as alterações. Você deve reenviar o suplemento com o manifesto de suplemento atualizado que inclui um novo número da versão. Você também deve se certificar de atualizar o número da versão do suplemento no formulário de envio para corresponder ao novo número da versão do manifesto. Para suplementos do Outlook, verifique se o elemento [Id](https://dev.office.com/reference/add-ins/manifest/id) contém um UUID diferente do manifesto de suplemento.
    
