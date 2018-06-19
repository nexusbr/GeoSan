Insere novas colunas para poder registrar o indicador de produtividade de cadastro de novas ligações em um ramal novo ou existente. Registra quem cadastrou e quando cadastrou.

alter table ramais_agua_ligacao add DATA_LOG varchar(45) null, USUARIO_LOG varchar(30) null

alter table RAMAIS_ESGOTO_LIGACAO add DATA_LOG varchar(45) null, USUARIO_LOG varchar(30) null