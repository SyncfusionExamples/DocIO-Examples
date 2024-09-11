#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM registry.access.redhat.com/ubi8/dotnet-80-runtime AS base
USER root
RUN microdnf install -y fontconfig \
    && microdnf clean all
WORKDIR /app

# Use Red Hat Universal Base Image (UBI) 8 with .NET 8.0 SDK for building the application
FROM registry.access.redhat.com/ubi8/dotnet-80 AS build
WORKDIR /src
COPY ["WordToPDFDockerSample.csproj", ""]
RUN dotnet restore "./WordToPDFDockerSample.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "WordToPDFDockerSample.csproj" -c Release -o /app/build

# Publish the application
FROM build AS publish
RUN dotnet publish "WordToPDFDockerSample.csproj" -c Release -o /app/publish

# Use the .NET 8.0 runtime to run the application
FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "WordToPDFDockerSample.dll"]
