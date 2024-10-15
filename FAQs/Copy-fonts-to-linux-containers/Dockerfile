#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/runtime:8.0-jammy AS base
RUN apt-get update -y && apt-get install fontconfig -y
WORKDIR /app
#Copy fonts to linux containers from a folder.
COPY ["Fonts/*.*", "/usr/local/share/fonts/"]
FROM mcr.microsoft.com/dotnet/sdk:8.0-jammy AS build
WORKDIR /src
COPY ["Copy-fonts-to-linux-containers.csproj", "."]
RUN dotnet restore "./Copy-fonts-to-linux-containers.csproj"
COPY . .
WORKDIR "/src/."
RUN dotnet build "Copy-fonts-to-linux-containers.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "Copy-fonts-to-linux-containers.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "Copy-fonts-to-linux-containers.dll"]