package mcp.development_guides.project;

import mcp.development_guides.project.infrastructure.excel.specialized.ExcelMCPService;
import org.springframework.ai.support.ToolCallbacks;
import org.springframework.ai.tool.ToolCallback;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

import java.util.List;

@SpringBootApplication
public class MCPApplication {

	public static void main(String[] args) {
		SpringApplication.run(MCPApplication.class, args);
	}

	@Bean
	public List<ToolCallback> iaTools(ExcelMCPService excelMCPService) {
		return List.of(ToolCallbacks.from(excelMCPService));
	}

}
