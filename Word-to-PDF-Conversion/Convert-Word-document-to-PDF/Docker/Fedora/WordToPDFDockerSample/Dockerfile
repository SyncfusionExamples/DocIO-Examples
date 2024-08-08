#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM fedora:latest

# Install .NET 8.0 SDK and runtime
Run dnf install dotnet-sdk-8.0 -y
RUN dnf install dotnet-runtime-8.0 -y

# Install fontconfig for font support
RUN dnf install fontconfig -y

# Set the working directory
WORKDIR /app

# Copy the application files
COPY ["WordToPDFDockerSample.csproj", "."]
RUN dotnet restore "WordToPDFDockerSample.csproj"
COPY . .

# Build the application
RUN dotnet build "WordToPDFDockerSample.csproj" -c Release -o /app/build

# Publish the application
RUN dotnet publish "WordToPDFDockerSample.csproj" -c Release -o /app/publish

# Set the final stage to use the runtime image
ENTRYPOINT ["dotnet", "WordToPDFDockerSample.dll"]
