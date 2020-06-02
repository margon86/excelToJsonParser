package com.exceltojason.convertirdor.util;

import java.text.Normalizer;

public class ParserCaracteresEspeciales {
	
	//Este metodo nos permite normalizar los carateres especiales como por ej ÁáÉéÍíÓóÚúÑñÜü
	public String parsearCaracteresEspeciales(String input) {
		String retorno = Normalizer
					        .normalize(input, Normalizer.Form.NFD)
					        .replaceAll("[^\\p{ASCII}]", "");
		return retorno;
	}

}
