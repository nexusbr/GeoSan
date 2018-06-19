USE [ArturNogueira-B]


/* insere parâmetros de bomba */

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(3, 1, 'POTÊNCIA', 'False', 0.0000, 0.0000, 0, 2, NULL)

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(3, 2, 'CARGA', 'False', 0.0000, 0.0000, 0, 2, NULL)

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(3, 3, 'VAZÃO', 'False', 0.0000, 0.0000, 0, 2, NULL)

/* insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(3, 4, 'REND_VAZÃO', 'False', 0.0000, 0.0000, 0, 2, NULL) */

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(3, 4, 'RENDIMENTO', 'False', 0.0000, 0.0000, 0, 2, NULL)

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(3, 5, 'CURVA DA BOMBA', 'False', 0.0000, 0.0000, 0, 2, NULL)

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(3, 6, 'CURVA DE REND', 'False', 0.0000, 0.0000, 0, 2, NULL)


/* Apaga todos os metadados de bomba */

delete from WaterComponentsSubTypes where id_Type = 3 and id_SubType = 1 and Description_ = 'POTÊNCIA'

delete from WaterComponentsSubTypes where id_Type = 3 and id_SubType = 2 and Description_ = 'CARGA'

delete from WaterComponentsSubTypes where id_Type = 3 and id_SubType = 3 and Description_ = 'VAZÃO'

delete from WaterComponentsSubTypes where id_Type = 3 and id_SubType = 4 and Description_ = 'RENDIMENTO'

delete from WaterComponentsSubTypes where id_Type = 3 and id_SubType = 5 and Description_ = 'CURVA DA BOMBA'

delete from WaterComponentsSubTypes where id_Type = 3 and id_SubType = 6 and Description_ = 'CURVA DE REND'

/* Apaga todos os dados cadastrados no GeoSan das bombas (TODOS) */

delete from WATERCOMPONENTSDATA where ID_TYPE = 3 and ID_SUBTYPE = 1

delete from WATERCOMPONENTSDATA where ID_TYPE = 3 and ID_SUBTYPE = 2

delete from WATERCOMPONENTSDATA where ID_TYPE = 3 and ID_SUBTYPE = 3

delete from WATERCOMPONENTSDATA where ID_TYPE = 3 and ID_SUBTYPE = 4

delete from WATERCOMPONENTSDATA where ID_TYPE = 3 and ID_SUBTYPE = 5

delete from WATERCOMPONENTSDATA where ID_TYPE = 3 and ID_SUBTYPE = 6

/* insere parâmetros de um RNV */

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(28, 1, 'ALT INICIAL', 'False', 0.0000, 0.0000, 0, 2, NULL)

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(28, 2, 'ALT MIN', 'False', 0.0000, 0.0000, 0, 2, NULL)

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(28, 3, 'ALT MAX', 'False', 0.0000, 0.0000, 0, 2, NULL)

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(28, 4, 'DIÂMETRO', 'False', 0.0000, 0.0000, 0, 2, NULL)

insert into WaterComponentsSubTypes (id_Type, id_SubType, Description_, Selection_, Max_, Min_, DefaultValue, DataType, EPAREF) values
(28, 5, 'CURVA DE VOLUME', 'False', 0.0000, 0.0000, 0, 2, NULL)

/* Apaga todos os metadados de RNV */

delete from WaterComponentsSubTypes where id_Type = 28 and id_SubType = 1 and Description_ = 'ALT INICIAL'

delete from WaterComponentsSubTypes where id_Type = 28 and id_SubType = 2 and Description_ = 'ALT MIN'

delete from WaterComponentsSubTypes where id_Type = 28 and id_SubType = 3 and Description_ = 'ALT MAX'

delete from WaterComponentsSubTypes where id_Type = 28 and id_SubType = 4 and Description_ = 'DIÂMETRO'

delete from WaterComponentsSubTypes where id_Type = 28 and id_SubType = 4 and Description_ = 'CURVA DE VOLUME'

/* Apaga todos os dados cadastrados no GeoSan das bombas (TODOS) */

delete from WATERCOMPONENTSDATA where ID_TYPE = 28 and ID_SUBTYPE = 1

delete from WATERCOMPONENTSDATA where ID_TYPE = 28 and ID_SUBTYPE = 2

delete from WATERCOMPONENTSDATA where ID_TYPE = 28 and ID_SUBTYPE = 3

delete from WATERCOMPONENTSDATA where ID_TYPE = 28 and ID_SUBTYPE = 4

delete from WATERCOMPONENTSDATA where ID_TYPE = 28 and ID_SUBTYPE = 5