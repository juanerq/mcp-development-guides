package mcp.development_guides.project;

import jakarta.annotation.PostConstruct;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.ai.tool.annotation.Tool;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class CourseService {
    private static final Logger logger = LoggerFactory.getLogger(CourseService.class);
    private final List<Course> courses = new ArrayList<>();

    @Tool(name = "dv_get_courses", description = "Retrieve a list of available courses")
    public List<Course> getCourses() {
        return courses;
    }

    @Tool(name = "dv_get_course_by_title", description = "Retrieve a course by its title")
    public Course getCourseByTitle(String title) {
        return courses.stream()
                .filter(course -> course.title().equalsIgnoreCase(title))
                .findFirst()
                .orElse(null);
    }


    @PostConstruct
    public void init() {
        logger.info("Initializing CourseService with default courses");
        courses.addAll(List.of(
                new Course("Spring Boot Basics", "https://example.com/spring-boot-basics"),
                new Course("Advanced Spring Boot", "https://example.com/advanced-spring-boot"),
                new Course("Microservices with Spring Boot", "https://example.com/microservices-spring-boot"),
                new Course("Spring Boot Testing", "https://example.com/spring-boot-testing"),
                new Course("Spring Boot Security", "https://example.com/spring-boot-security")
        ));
    }
}
