# Use the official .NET 8.0 SDK image from Microsoft for building the application
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
RUN apt-get update -y && apt-get install libfontconfig -y
WORKDIR /src

# Copy the project file and restore dependencies
COPY ["WordToPDFDockerSample.csproj", "."]
RUN dotnet restore "WordToPDFDockerSample.csproj"

# Copy the rest of the application code and build the application
COPY . .
RUN dotnet build "WordToPDFDockerSample.csproj" -c Release -o /app/build

# Publish the application
FROM build AS publish
RUN dotnet publish "WordToPDFDockerSample.csproj" -c Release -o /app/publish /p:UseAppHost=false

# Use CentOS 8 as the base image for the final runtime
FROM centos:8

# Install .NET 8.0 runtime
RUN dnf install -y https://packages.microsoft.com/config/centos/8/prod.repo \
    && dnf install -y dotnet-runtime-8.0 \
    && dnf clean all

# Install fontconfig for fonts
RUN dnf install -y fontconfig

# Set the working directory
WORKDIR /app

# Copy the published application from the previous stage
COPY --from=publish /app/publish .

# Set the entry point for the container
ENTRYPOINT ["dotnet", "WordToPDFDockerSample.dll"]
