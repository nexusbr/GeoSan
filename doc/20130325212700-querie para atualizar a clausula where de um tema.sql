update [valinhosgeosan].[dbo].[te_theme] 
set [generate_attribute_where] = 'object_id in (select object_id_ from WATERCOMPONENTS WHERE ID_TYPE=' + CHAR (39) + '27' + CHAR(39) + ')'
where [theme_id] = 1454