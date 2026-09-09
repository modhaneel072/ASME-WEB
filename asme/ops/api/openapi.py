"""OpenAPI 3.1 document assembled from the blueprint's routes and their pydantic contracts."""

from __future__ import annotations

from flask import current_app, jsonify

from asme.ops.api import bp

ERROR_SCHEMA = {
    "type": "object",
    "properties": {
        "ok": {"type": "boolean", "const": False},
        "code": {"type": "string"},
        "error": {"type": "string"},
        "details": {"type": "array", "items": {"type": "object", "properties": {"field": {"type": ["string", "null"]}, "message": {"type": "string"}}}},
    },
    "required": ["ok", "code", "error"],
}


def build_document() -> dict:
    components: dict = {"schemas": {"Error": ERROR_SCHEMA}}
    paths: dict = {}
    for rule in current_app.url_map.iter_rules():
        if not rule.endpoint.startswith("ops_api."):
            continue
        view = current_app.view_functions[rule.endpoint]
        meta = getattr(view, "_contract", {}) or {}
        path = rule.rule
        if path.startswith("/api/v1"):
            path = path[len("/api/v1"):] or "/"
        for arg in rule.arguments:
            path = path.replace(f"<int:{arg}>", f"{{{arg}}}").replace(f"<{arg}>", f"{{{arg}}}")
        endpoint_name = rule.endpoint.split(".", 1)[1]
        doc_summary = (view.__doc__ or "").strip().splitlines()[0] if view.__doc__ else ""
        for method in sorted(rule.methods - {"HEAD", "OPTIONS"}):
            operation = {
                "operationId": endpoint_name + ("" if len(rule.methods - {"HEAD", "OPTIONS"}) == 1 else f"_{method.lower()}"),
                "summary": meta.get("summary") or doc_summary or endpoint_name.replace("_", " ").capitalize(),
                "tags": meta.get("tags") or [path.split("/")[3] if len(path.split("/")) > 3 else "ops"],
                "parameters": [{"name": arg, "in": "path", "required": True, "schema": {"type": "integer" if f"<int:{arg}>" in rule.rule else "string"}} for arg in rule.arguments],
                "responses": {
                    "200": {"description": "OK", "content": {"application/json": {"schema": {"type": "object", "properties": {"ok": {"type": "boolean"}, "payload": {}}}}}},
                    "400": {"description": "Validation error", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/Error"}}}},
                    "401": {"description": "Login required", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/Error"}}}},
                    "403": {"description": "Forbidden (includes `permission`)", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/Error"}}}},
                    "404": {"description": "Not found (also returned for records in another organisation)", "content": {"application/json": {"schema": {"$ref": "#/components/schemas/Error"}}}},
                },
            }
            req_model = meta.get("request")
            if req_model is not None and method in {"POST", "PATCH", "PUT"}:
                schema = req_model.model_json_schema(ref_template="#/components/schemas/{model}")
                defs = schema.pop("$defs", {})
                components["schemas"].update(defs)
                components["schemas"][req_model.__name__] = schema
                operation["requestBody"] = {"required": True, "content": {"application/json": {"schema": {"$ref": f"#/components/schemas/{req_model.__name__}"}}}}
            resp_model = meta.get("response")
            if resp_model is not None:
                schema = resp_model.model_json_schema(ref_template="#/components/schemas/{model}")
                defs = schema.pop("$defs", {})
                components["schemas"].update(defs)
                components["schemas"][resp_model.__name__] = schema
                operation["responses"]["200"]["content"]["application/json"]["schema"]["properties"]["payload"] = {"$ref": f"#/components/schemas/{resp_model.__name__}"}
            paths.setdefault(path, {})[method.lower()] = operation
    return {
        "openapi": "3.1.0",
        "info": {"title": "ASME Ops API", "version": "1.0.0", "description": "Organisation-scoped operations API. Send `X-Requested-With: ASME-Ops` on POST/PATCH/PUT/DELETE. List endpoints accept filter[...], sort, q, page[limit], page[cursor]."},
        "servers": [{"url": "/api/v1"}],
        "paths": dict(sorted(paths.items())),
        "components": components,
    }


@bp.get("/ops/openapi.json")
def openapi():
    return jsonify(build_document())
