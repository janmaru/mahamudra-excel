# .NET Core Project Assistant

You are assisting with a C# .NET Core project. Follow these guidelines:

## Code Style
- Use PascalCase for public members, classes, and methods
- Use camelCase for private fields (prefix with underscore: _fieldName)
- Use meaningful, descriptive names
- Prefer async/await for I/O operations
- Use nullable reference types where appropriate
- Follow SOLID principles

## Project Structure
- Controllers go in `/Controllers`
- Services go in `/Services` with interfaces in `/Services/Interfaces`
- Models/DTOs go in `/Models` or `/DTOs`
- Data access goes in `/Data` or `/Repositories`
- Configuration goes in `/Configuration`

## Best Practices
- Use dependency injection for all services
- Implement interfaces for testability
- Use ILogger<T> for logging
- Handle exceptions with proper try-catch or middleware
- Validate input with Data Annotations or FluentValidation
- Use Entity Framework Core with migrations for data access
- Return appropriate HTTP status codes from API endpoints

## Common Commands Reference
- `dotnet build` - Build the project
- `dotnet run` - Run the application
- `dotnet test` - Run unit tests
- `dotnet ef migrations add <Name>` - Add EF migration
- `dotnet ef database update` - Apply migrations

## When Writing Code
1. Check existing patterns in the codebase first
2. Reuse existing services and utilities
3. Add XML documentation for public APIs
4. Consider edge cases and null checks
5. Write unit tests for new functionality

## Response Format
- Explain your reasoning briefly
- Provide complete, working code
- Include necessary using statements
- Note any NuGet packages required
