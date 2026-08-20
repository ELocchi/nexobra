# Nexobra — Diálogo Engenharia

Plataforma Nexobra configurada para a Diálogo Engenharia e atualmente aplicada à obra BoulevarDiálogo Butantã, com:

- controle por torre, pavimento, apartamento e serviço;
- ocorrências com fotos e histórico em PDF;
- dashboards e evolução visual dos pavimentos;
- placas individuais ou em lote, com QR Code de visitante por unidade;
- persistência em SQLite;
- acesso administrativo protegido por login.

## Instalação

É necessário ter Python 3 instalado.

```bash
python3 -m pip install -r requirements.txt
cp configuracao.exemplo.json configuracao.local.json
```

Edite `configuracao.local.json` e defina o usuário, a senha do engenheiro e um
`segredo_visitante` longo. Mantenha esse segredo: trocá-lo invalida os QR Codes já
impressos. Esse arquivo e o banco SQLite são ignorados pelo Git.

## Executar

No macOS, abra `INICIAR-PAINEL.command`, ou execute:

```bash
python3 servidor.py
```

No computador, acesse `http://127.0.0.1:8000`.

Para acesso por celular, use o endereço de rede mostrado no terminal e mantenha o celular na mesma rede Wi-Fi.

## Dados locais

O arquivo `acompanhamento.db` é criado automaticamente e não deve ser publicado, pois pode conter dados e fotos da obra.

## Publicação no Render

O projeto inclui um `render.yaml` para criar um serviço web com disco persistente. No Render:

1. conecte o repositório privado do GitHub;
2. crie um Blueprint usando este repositório;
3. informe `OBRA_USUARIO` e `OBRA_SENHA` quando solicitado;
4. confirme o serviço Starter e o disco persistente de 1 GB.

O serviço usará a porta definida pelo ambiente e guardará o SQLite em
`/var/data/acompanhamento.db`. O Render fornecerá um endereço HTTPS público.
