#Nota: queries feitas e testadas em MySQL

SELECT * FROM montadora;
SELECT * FROM produto;
SELECT * FROM produto_veiculo;
SELECT * FROM veiculo;

# 1 - SELECT: Todos os "Produtos" (Codigo e Descrição), e "Montadoras" (Descrição) / "Veiculos" (Descrição) / "ObsProdutoVeiculo" correspondentes - Ordenação: Primeiro o Veiculo "MONZA", depois o "FOCUS" e o restante em Ordem Alfabética por Descrição de "Produto, Montadora, Veiculo".
	SELECT p.CodigoProduto AS 'Código', p.DescricaoProduto AS 'Produto', m.DescricaoMontadora AS 'Montadora', v.DescricaoVeiculo AS 'Veículo', pv.`ObsProdutoVeiculo` AS 'OBS'
	  FROM produto P
	  JOIN produto_veiculo PV ON (p.CodigoProduto = pv.CodigoProduto)
	  JOIN veiculo V ON (pv.CodigoVeiculo = v.CodigoVeiculo)
	  JOIN montadora M ON (v.CodigoMontadora = m.CodigoMontadora)
	 ORDER BY FIELD(v.DescricaoVeiculo, 'FOCUS', 'MONZA') DESC, p.DescricaoProduto, m.DescricaoMontadora, v.DescricaoVeiculo;


# 2 - SELECT: "Produto" (Codigo e Descrição), "Veiculo" (Descrição) - Apenas para Veiculos Sem Montadora.
	SELECT p.CodigoProduto AS 'Código', p.DescricaoProduto AS 'Produto', v.DescricaoVeiculo AS 'Veículo'
	  FROM produto P
	  JOIN produto_veiculo PV ON (p.CodigoProduto = pv.CodigoProduto)
	  JOIN veiculo V ON (pv.CodigoVeiculo = v.CodigoVeiculo)
	 WHERE v.CodigoMontadora IS NULL;
	 
# 3 - SELECT: "Montadoras" (Descrição) e a "Quantidade de Produtos" para cada Montadora.
	SELECT m.DescricaoMontadora AS 'Montadora', COUNT(*) AS 'Quantidade'
	  FROM produto_veiculo PV 
	  JOIN veiculo V ON (pv.CodigoVeiculo = v.CodigoVeiculo)
	  JOIN montadora M ON (v.CodigoMontadora = m.CodigoMontadora)
	 GROUP BY m.CodigoMontadora;