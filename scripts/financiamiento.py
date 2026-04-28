from __future__ import annotations

from datetime import datetime
import re
import unicodedata


def normalize_text_key(value: str) -> str:
    clean = (value or "").strip().lower()
    if not clean:
        return ""
    clean = unicodedata.normalize("NFKD", clean)
    clean = "".join(char for char in clean if not unicodedata.combining(char))
    return re.sub(r"\s+", " ", clean).strip()


def normalize_si_no(value: str, *, default: str = "No") -> str:
    key = normalize_text_key(str(value or ""))
    if key in {"si", "sí", "s", "yes", "true", "1", "ramo 33", "ramo33", "ramo 28", "ramo28"}:
        return "Si"
    if key in {"no", "n", "false", "0", ""}:
        return default
    return default


def is_remanente_fuente(value: str) -> bool:
    key = normalize_text_key(value)
    return (
        key.startswith("remanente")
        or key.startswith("remanentes")
        or key.startswith("rea:")
        or key.startswith("seguimiento")
    )


def normalize_origen_fuente(value: str) -> str:
    clean = " ".join((value or "").split())
    key = normalize_text_key(clean)
    if key in {"remanente", "remanentes"}:
        return "Remanentes"
    if key in {"del ejercicio", "ejercicio", "del_ejercicio"}:
        return "Del Ejercicio"
    return clean


def infer_origen_fuente(fuente_nombre: str, origen_fuente: str = "") -> str:
    origen = normalize_origen_fuente(origen_fuente)
    if origen in {"Del Ejercicio", "Remanentes"}:
        return origen
    return "Remanentes" if is_remanente_fuente(fuente_nombre) else "Del Ejercicio"


def origen_fuente_sql(alias: str = "") -> str:
    prefix = f"{alias}." if alias else ""
    fuente_col = f"{prefix}fuente_financiamiento"
    origen_col = f"{prefix}origen_fuente"
    fuente_key = f"LOWER(TRIM(COALESCE({fuente_col}, '')))"
    origen_key = f"LOWER(TRIM(COALESCE({origen_col}, '')))"
    return (
        "CASE "
        f"WHEN {origen_key} IN ('remanente', 'remanentes') THEN 'Remanentes' "
        f"WHEN {origen_key} IN ('del ejercicio', 'ejercicio', 'del_ejercicio') THEN 'Del Ejercicio' "
        f"WHEN {fuente_key} LIKE 'remanente%' THEN 'Remanentes' "
        f"WHEN {fuente_key} LIKE 'rea:%' THEN 'Remanentes' "
        f"WHEN {fuente_key} LIKE 'seguimiento%' THEN 'Remanentes' "
        "ELSE 'Del Ejercicio' END"
    )


def resolve_fuente_catalogo(
    db,
    fuente_nombre: str,
    *,
    normalizer,
    create_missing: bool = False,
) -> tuple[int | None, str]:
    clean_name = normalizer(fuente_nombre)
    if not clean_name:
        raise ValueError("Debes escribir la nueva fuente.")
    row = db.execute(
        """
        SELECT id, TRIM(COALESCE(nombre, '')) AS nombre
        FROM fuentes_financiamiento
        WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))
        LIMIT 1
        """,
        (clean_name,),
    ).fetchone()
    if row:
        return int(row["id"]), (row["nombre"] or "").strip()
    if not create_missing:
        return None, clean_name
    cursor = db.execute(
        """
        INSERT INTO fuentes_financiamiento (nombre, ramo_33, ramo_28, origen_fuente, created_at)
        VALUES (?, ?, ?, ?, ?)
        """,
        (
            clean_name,
            "No",
            "No",
            infer_origen_fuente(clean_name, ""),
            datetime.now().strftime("%Y-%m-%d %H:%M"),
        ),
    )
    return int(cursor.lastrowid), clean_name


def get_fuente_clasificacion(
    db,
    fuente_nombre: str,
    *,
    fuente_id: int | None = None,
    normalizer,
    si_no_normalizer=normalize_si_no,
    origen_normalizer=infer_origen_fuente,
) -> dict:
    row = None
    if fuente_id is not None and fuente_id >= 0:
        row = db.execute(
            """
            SELECT
                TRIM(COALESCE(nombre, '')) AS nombre,
                TRIM(COALESCE(ramo_33, 'No')) AS ramo_33,
                TRIM(COALESCE(ramo_28, 'No')) AS ramo_28,
                TRIM(COALESCE(origen_fuente, '')) AS origen_fuente
            FROM fuentes_financiamiento
            WHERE id = ?
            LIMIT 1
            """,
            (fuente_id,),
        ).fetchone()
    if row is None:
        clean_name = normalizer(fuente_nombre)
        if clean_name:
            row = db.execute(
                """
                SELECT
                    TRIM(COALESCE(nombre, '')) AS nombre,
                    TRIM(COALESCE(ramo_33, 'No')) AS ramo_33,
                    TRIM(COALESCE(ramo_28, 'No')) AS ramo_28,
                    TRIM(COALESCE(origen_fuente, '')) AS origen_fuente
                FROM fuentes_financiamiento
                WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))
                LIMIT 1
                """,
                (clean_name,),
            ).fetchone()
    source_name = (row["nombre"] if row else fuente_nombre) or ""
    return {
        "ramo_33": si_no_normalizer(row["ramo_33"] if row else "No"),
        "ramo_28": si_no_normalizer(row["ramo_28"] if row else "No"),
        "origen_fuente": origen_normalizer(source_name, row["origen_fuente"] if row else ""),
    }


def list_fuentes_financiamiento_admin(
    db,
    ejercicio: str,
    *,
    normalizer,
    si_no_normalizer=normalize_si_no,
    origen_normalizer=infer_origen_fuente,
) -> list[dict]:
    ejercicio_clean = " ".join((ejercicio or "").split())
    sources: dict[str, dict] = {}

    def safe_source_id(raw_value) -> int | None:
        try:
            source_id = int(raw_value)
        except (TypeError, ValueError):
            return None
        return source_id if source_id > 0 else None

    def ensure_source(
        nombre: str,
        *,
        source_id=None,
        ramo_33: str = "",
        ramo_28: str = "",
        origen_fuente: str = "",
        catalogo: bool = False,
    ) -> dict | None:
        canonical_name = normalizer(nombre)
        if not canonical_name:
            return None
        key = normalize_text_key(canonical_name)
        if not key:
            return None
        source_id_int = safe_source_id(source_id)
        item = sources.get(key)
        if item is None:
            item = {
                "id": source_id_int or "",
                "nombre": canonical_name,
                "ramo_33": si_no_normalizer(ramo_33 or "No"),
                "ramo_28": si_no_normalizer(ramo_28 or "No"),
                "origen_fuente": origen_normalizer(canonical_name, origen_fuente or ""),
                "catalogo": bool(catalogo),
                "cargas_count": 0,
                "cargas_observaciones": 0,
                "observaciones_count": 0,
                "entes_count": 0,
            }
            sources[key] = item
        if source_id_int and not item.get("id"):
            item["id"] = source_id_int
        if catalogo:
            item["catalogo"] = True
            item["nombre"] = canonical_name
            item["ramo_33"] = si_no_normalizer(ramo_33 or "No")
            item["ramo_28"] = si_no_normalizer(ramo_28 or "No")
            item["origen_fuente"] = origen_normalizer(canonical_name, origen_fuente or "")
        return item

    catalog_rows = db.execute(
        """
        SELECT
            id,
            TRIM(COALESCE(nombre, '')) AS nombre,
            TRIM(COALESCE(ramo_33, 'No')) AS ramo_33,
            TRIM(COALESCE(ramo_28, 'No')) AS ramo_28,
            TRIM(COALESCE(origen_fuente, '')) AS origen_fuente
        FROM fuentes_financiamiento
        WHERE TRIM(COALESCE(nombre, '')) != ''
        ORDER BY nombre ASC
        """
    ).fetchall()
    for row in catalog_rows:
        ensure_source(
            row["nombre"],
            source_id=row["id"],
            ramo_33=row["ramo_33"],
            ramo_28=row["ramo_28"],
            origen_fuente=row["origen_fuente"],
            catalogo=True,
        )

    if ejercicio_clean:
        cargas_rows = db.execute(
            """
            SELECT
                cm.fuente_id,
                COALESCE(NULLIF(TRIM(cm.fuente_nombre), ''), TRIM(ff.nombre), '') AS nombre,
                COALESCE(cm.cantidad_sa, 0) AS cantidad_sa,
                COALESCE(cm.cantidad_pdp, 0) AS cantidad_pdp,
                COALESCE(cm.cantidad_pras, 0) AS cantidad_pras,
                COALESCE(cm.cantidad_pefcf, 0) AS cantidad_pefcf,
                COALESCE(cm.cantidad_r, 0) AS cantidad_r
            FROM cargas_manuales AS cm
            LEFT JOIN fuentes_financiamiento AS ff
              ON ff.id = cm.fuente_id
            WHERE TRIM(COALESCE(cm.ejercicio, '')) = ?
            """,
            (ejercicio_clean,),
        ).fetchall()
        for row in cargas_rows:
            item = ensure_source(row["nombre"], source_id=row["fuente_id"])
            if item is None:
                continue
            item["cargas_count"] += 1
            item["cargas_observaciones"] += sum(
                int(row[field] or 0)
                for field in (
                    "cantidad_sa",
                    "cantidad_pdp",
                    "cantidad_pras",
                    "cantidad_pefcf",
                    "cantidad_r",
                )
            )

        observaciones_rows = db.execute(
            """
            SELECT
                TRIM(COALESCE(fuente_financiamiento, '')) AS nombre,
                COUNT(*) AS total
            FROM observaciones
            WHERE TRIM(COALESCE(ejercicio, '')) = ?
              AND TRIM(COALESCE(fuente_financiamiento, '')) != ''
            GROUP BY LOWER(TRIM(COALESCE(fuente_financiamiento, '')))
            """,
            (ejercicio_clean,),
        ).fetchall()
        for row in observaciones_rows:
            item = ensure_source(row["nombre"])
            if item is None:
                continue
            item["observaciones_count"] += int(row["total"] or 0)

        entes_rows = db.execute(
            """
            SELECT
                ff.id,
                TRIM(COALESCE(ff.nombre, '')) AS nombre,
                COUNT(*) AS total
            FROM entes_fuentes AS ef
            JOIN fuentes_financiamiento AS ff
              ON ff.id = ef.fuente_id
            WHERE TRIM(COALESCE(ef.ejercicio, '')) = ?
            GROUP BY ff.id, LOWER(TRIM(COALESCE(ff.nombre, '')))
            """,
            (ejercicio_clean,),
        ).fetchall()
        for row in entes_rows:
            item = ensure_source(row["nombre"], source_id=row["id"])
            if item is None:
                continue
            item["entes_count"] += int(row["total"] or 0)

    for item in sources.values():
        item["total_uso"] = (
            int(item["cargas_count"] or 0)
            + int(item["observaciones_count"] or 0)
            + int(item["entes_count"] or 0)
        )

    return sorted(
        sources.values(),
        key=lambda item: (
            0 if int(item.get("total_uso") or 0) > 0 else 1,
            (item.get("nombre") or "").casefold(),
        ),
    )


def update_fuente_clasificacion(
    db,
    *,
    fuente_nombre: str,
    normalizer,
    fuente_id: int | None = None,
    ejercicio: str = "",
    ramo_33: str = "No",
    ramo_28: str = "No",
    origen_fuente: str = "",
) -> dict:
    source_id = fuente_id
    source_name = normalizer(fuente_nombre)
    if source_id is not None:
        source_row = db.execute(
            """
            SELECT id, TRIM(COALESCE(nombre, '')) AS nombre
            FROM fuentes_financiamiento
            WHERE id = ?
            LIMIT 1
            """,
            (source_id,),
        ).fetchone()
        if not source_row:
            raise ValueError("La fuente seleccionada no existe.")
        source_name = source_name or normalizer(source_row["nombre"] or "")
    if not source_name:
        raise ValueError("Debes seleccionar una fuente.")
    if source_id is None:
        source_id, source_name = resolve_fuente_catalogo(
            db,
            source_name,
            normalizer=normalizer,
            create_missing=True,
        )
    if source_id is None:
        raise ValueError("No se pudo guardar la fuente.")

    ramo_33_clean = normalize_si_no(ramo_33)
    ramo_28_clean = normalize_si_no(ramo_28)
    origen_clean = infer_origen_fuente(source_name, origen_fuente)

    db.execute(
        """
        UPDATE fuentes_financiamiento
        SET ramo_33 = ?,
            ramo_28 = ?,
            origen_fuente = ?
        WHERE id = ?
        """,
        (ramo_33_clean, ramo_28_clean, origen_clean, source_id),
    )

    ejercicio_clean = " ".join((ejercicio or "").split())
    cargas_where = [
        "(fuente_id = ? OR LOWER(TRIM(COALESCE(fuente_nombre, ''))) = LOWER(TRIM(?)))"
    ]
    cargas_params: list[object] = [source_id, source_name]
    if ejercicio_clean:
        cargas_where.append("TRIM(COALESCE(ejercicio, '')) = ?")
        cargas_params.append(ejercicio_clean)
    cargas_cursor = db.execute(
        f"""
        UPDATE cargas_manuales
        SET ramo_33 = ?,
            ramo_28 = ?,
            origen_fuente = ?
        WHERE {" AND ".join(cargas_where)}
        """,
        [ramo_33_clean, ramo_28_clean, origen_clean, *cargas_params],
    )

    obs_where = [
        "LOWER(TRIM(COALESCE(fuente_financiamiento, ''))) = LOWER(TRIM(?))"
    ]
    obs_params: list[object] = [source_name]
    if ejercicio_clean:
        obs_where.append("TRIM(COALESCE(ejercicio, '')) = ?")
        obs_params.append(ejercicio_clean)
    observaciones_cursor = db.execute(
        f"""
        UPDATE observaciones
        SET ramo_33 = ?,
            ramo_28 = ?,
            origen_fuente = ?
        WHERE {" AND ".join(obs_where)}
        """,
        [ramo_33_clean, ramo_28_clean, origen_clean, *obs_params],
    )

    return {
        "fuente_id": source_id,
        "fuente_nombre": source_name,
        "ramo_33": ramo_33_clean,
        "ramo_28": ramo_28_clean,
        "origen_fuente": origen_clean,
        "cargas_actualizadas": max(cargas_cursor.rowcount or 0, 0),
        "observaciones_actualizadas": max(observaciones_cursor.rowcount or 0, 0),
    }
