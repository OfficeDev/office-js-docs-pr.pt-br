
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>Realizar o sideload de suplementos do Office em um iPad ou Mac para teste

Para ver como seu suplemento será executado no Office para iOS, você pode realizar o sideload do manifesto do seu suplemento em um iPad usando o iTunes, ou realizar o sideload do manifesto do suplemento diretamente no Office para Mac. Esta ação não permite definir pontos de interrupção e depurar o código do seu suplemento enquanto ele estiver em execução, mas é possível ver como ele se comporta e verificar se a interface do usuário é utilizável e está sendo processada adequadamente. 

## <a name="prerequisites-for-office-for-ios"></a>Pré-requisitos do Office para iOS



- Um computador com Windows ou Mac com [iTunes](http://www.apple.com/itunes/download/) instalado.
    
- Um iPad executando o iOS 8.2 ou posterior com [Excel para iPad](https://itunes.apple.com/us/app/microsoft-excel/id586683407?mt=8) instalado e um cabo de sincronização.
    
- O arquivo de manifesto .xml para o suplemento que você deseja testar.
    

## <a name="prerequisites-for-office-for-mac"></a>Pré-requisitos do Office para Mac



- Um Mac executando o OS X v10.10 "Yosemite" ou posterior com [Office para Mac](https://products.office.com/en-us/buy/compare-microsoft-office-products?tab=omac) instalado.
    
- Word para Mac versão 15.18 (160109).
   
- Excel para Mac versão 15.19 (160206).

- PowerPoint para Mac versão 15.24 (160614)
    
- O arquivo de manifesto .xml para o suplemento que você deseja testar.
    

## <a name="sideload-an-add-in-on-excel-or-word-for-ipad"></a>Realizar um sideload de um suplemento no Excel ou no Word para iPad

1. Use um cabo de sincronização para conectar seu iPad ao computador. Se estiver conectando o iPad ao computador pela primeira vez, será solicitado a responder **Confiar Neste Computador?** Escolha **Confiar** para continuar.

2. No iTunes, escolha o ícone do **iPad** abaixo da barra de menus.
    
    ![O ícone do iPad no iTunes](../../images/4ea35904-252e-45b4-88ad-14840d502bad.png)

3. Em **Ajustes** no lado esquerdo do iTunes, escolha **Aplicativos**.
    
    ![Configurações de aplicativos do iTunes](../../images/a12d1bb6-b39f-496b-83de-6ac00b0b97a5.png)

4. No lado direito do iTunes, role para baixo até **Compartilhamento de Arquivos**, e escolha **Excel** ou **Word** na coluna **Aplicativos**.
    
    ![Compartilhamento de arquivos do iTunes](../../images/3b2a53a2-e164-4ff0-ba42-83a8dc1a069f.png)

5. Na parte inferior da coluna Documentos do **Excel** ou do **Word**, escolha **Adicionar Arquivo** e selecione o arquivo de manifesto .xml do suplemento para o qual você deseja realizar sideload. 
    
6. Abra o aplicativo Excel ou Word em seu iPad. Se já estiver executando o aplicativo Excel ou Word, escolha o botão **Início**, feche e reinicie o aplicativo.
    
7. Abra um documento.
    
8. Escolha **Suplementos** na guia **Inserir**. O suplemento com sideload está disponível para inserção no cabeçalho **Desenvolvedor** na interface de usuário **Suplementos**.
    
    ![Inserir Suplementos no aplicativo do Excel](../../images/ed6033b0-ecec-4853-8ee7-9ef0884cb237.PNG)


## <a name="sideload-an-add-in-on-office-for-mac"></a>Realizar sideload de um suplemento no Office para Mac

> **Observação:** Para realizar o sideload de um suplemento do Outlook para Mac 2016, confira [Realizar sideload de suplementos do Outlook para teste](sideload-outlook-add-ins-for-testing.md).

1. Abra o **Terminal** e navegue até uma das pastas a seguir, onde você salvará o arquivo de manifesto do suplemento. Se a pasta `wef` não existir em seu computador, crie-a.
    
    - Para o Word:  `/Users/<username>/Library/Containers/com.microsoft.Word/Data/documents/wef`    
    - Para o Excel:  `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/documents/wef`
    - Para o PowerPoint: `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/documents/wef`
    
2. Abra a pasta no **Finder** usando o comando `open .` (incluindo o ponto final). Copie o arquivo de manifesto do suplemento nessa pasta.
    
    ![Pasta Wef no Office para Mac](../../images/bca689f8-bff4-421d-bc36-92c8ae0ddfba.png)

3. Abra o Word e abra um documento. Reinicie o Word se já estiver em execução.
    
4. No Word, escolha **Inserir** > **Suplementos** > **Meus Suplementos** (menu suspenso) e escolha seu suplemento.
    
    ![Meus Suplementos no Office para Mac](../../images/4593430c-b33e-4895-b2be-63fe3c4d08bc.png)

  > **Importante:** aplicativos em que foi feito o sideload não aparecerão na caixa de diálogo Meus Suplementos. Eles ficam visíveis apenas dentro do menu suspenso (pequena seta para baixo à direita de Meus Suplementos na guia **Inserir**). Os suplementos em que foi feito o sideload são exibidos na lista sob o título **Suplementos do Desenvolvedor** nesse menu. 
    
5. Verifique se o seu suplemento é exibido no Word.
    
    ![Suplemento do Office mostrado no Office para Mac](../../images/a5cb2efc-1180-45b4-85a6-13df817b9d2c.png)
    
> **Observação:** Os Suplementos muitas vezes são armazenados em cache no Office para Mac por questão de desempenho. Se você precisar forçar um recarregamento do seu suplemento enquanto estiver desenvolvendo-o, limpe a pasta Users/<usr>/Library/Containers/com.Microsoft.OsfWebHost/Data/. 

## <a name="additional-resources"></a>Recursos adicionais


- [Depurar suplementos do Office no iPad e no Mac](../testing/debug-office-add-ins-on-ipad-and-mac.md)
    
