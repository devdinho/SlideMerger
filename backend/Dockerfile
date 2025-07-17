# Etapa 1 - build
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /app

COPY . ./
RUN dotnet restore
RUN dotnet publish -c Release -o out

# Etapa 2 - runtime
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS runtime
WORKDIR /app

COPY --from=build /app/out ./
ENV ASPNETCORE_URLS=http://+:5030
EXPOSE 5030

ENTRYPOINT ["dotnet", "SlideMergerAPINew.dll"]
