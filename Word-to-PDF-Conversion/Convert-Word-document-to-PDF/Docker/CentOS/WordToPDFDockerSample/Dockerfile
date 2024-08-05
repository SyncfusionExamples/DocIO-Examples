#See https://aka.ms/containerfastmode to understand how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM centos:8

Run dnf install dotnet-sdk-8.0 -y
RUN dnf install dotnet-runtime-8.0 -y

RUN dnf install fontconfig -y

ENTRYPOINT ["dotnet" "WordToPDFDockerSample.dll"]