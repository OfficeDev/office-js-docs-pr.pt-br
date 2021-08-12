> [!TIP]
> Se você estiver testando seu suplemento em vários ambientes (por exemplo, em desenvolvimento, preparo, demonstração, etc.), recomendamos que você mantenha um arquivo de manifesto XML diferente para cada ambiente. Em cada arquivo de manifesto, você pode:
> - Especificar as URLs que correspondem ao ambiente.
> - Personalize valores de metadados como `DisplayName` e rótulos em `Resources` para indicar o ambiente, assim os usuários finais poderão identificar o ambiente correspondente de um suplemento por sideloaded. 
> - Personalize o `namespace` de funções personalizadas para indicar o ambiente, se o suplemento definir funções personalizadas.
> 
> Seguindo essas diretrizes, você simplificará o processo de teste e evitará problemas que, de outra forma, ocorreriam quando um suplemento fosse carregado simultaneamente em vários ambientes.