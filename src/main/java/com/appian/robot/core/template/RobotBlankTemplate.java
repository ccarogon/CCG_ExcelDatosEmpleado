package com.appian.robot.core.template;

import com.novayre.jidoka.client.api.IJidokaServer;
import com.novayre.jidoka.client.api.IRobot;
import com.novayre.jidoka.client.api.JidokaFactory;
import com.novayre.jidoka.client.api.annotations.Robot;
import com.novayre.jidoka.client.api.multios.IClient;
import com.novayre.jidoka.client.lowcode.IRobotVariable;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.io.FileOutputStream;
import java.io.File;

/**
 * The Class RobotBlankTemplate.
 */
@Robot
public class RobotBlankTemplate implements IRobot {

	/** The server. */
	private IJidokaServer< ? > server;
	
	/** The client. */
	private IClient client;

	/** variables de entrada **/
	private Map<String, IRobotVariable> variables;
	
	/**
	 * Initialize the modules
	 */
	public void start() {
		
		server = JidokaFactory.getServer();
		client = IClient.getInstance(this);

		server.debug("Robot inicializado");

		/**
		 * Variables de proceso
		 */

		variables = server.getWorkflowVariables();
		
	}
	/**
	 * Metodo Excel
	 */
	public void crearExcel() {

		IRobotVariable rvn = variables.get("Nombre");
		IRobotVariable rva = variables.get("Apellidos");


		//Libro en blanco
		HSSFWorkbook workbook = new HSSFWorkbook();

		//Crear una pestaña en blanco
		HSSFSheet sheet = workbook.createSheet("Datos empleado");

		//This data needs to be written (Object[])
		System.out.println("Insertamos datos en el fichero Excel");
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"ID", "NOMBRE", "APELLIDOS"});
		//data.put("2", new Object[] {1, rvn, rva});
		data.put("3", new Object[] {2, "Juan", "Lopez Muñoz"});
		data.put("4", new Object[] {3, "David", "Jaen Perez"});
		data.put("5", new Object[] {4, "Lucia", "Sanchez Perez"});
		System.out.println("Datos Excel en memoria");
		//Interaccionar con los datos y escribir en la pestaña
		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset)
		{
			Row row = sheet.createRow(rownum++);
			Object [] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr)
			{
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof String)
					cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
					cell.setCellValue((Integer)obj);
				System.out.println("Insertamos datos");
			}
		}
		try
		{
			//Escribir el libro en el fichero
			FileOutputStream out = new FileOutputStream(new File("C:\\Users\\ccarogon\\Documents\\RPA\\Result\\Datos_empleados.xls"));

			workbook.write(out);
			out.close();
			System.out.println("Datos_empleados.xls escrito satisfactoriamente en disco.");
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	/**
	 * End.
	 */
	public void end() {
		
	}
}
