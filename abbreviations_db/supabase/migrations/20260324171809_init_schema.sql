create extension if not exists pgcrypto;

create table public.abbreviations (

    uuid uuid primary key default gen_random_uuid(),

    data_source text not null,

    abbreviation text not null,

    created_at timestamptz not null default now(),

    constraint abbreviation_not_blank
        check (length(trim(abbreviation)) > 0)
);


create table public.expansions (

    id bigint generated always as identity primary key,

    expansion_text text not null,

    expansion_type text not null,

    created_at timestamptz not null default now(),

    constraint expansion_type_check
        check (expansion_type in ('historical', 'modern')),

    constraint unique_expansion
        unique (expansion_text, expansion_type)
);


create table public.images (

    id bigint generated always as identity primary key,

    abbreviation_uuid uuid not null
        references public.abbreviations(uuid)
        on delete cascade,

    image_path text not null,

    image_source text not null default 'upload',

    status text not null default 'pending',

    uploaded_by text,

    uploaded_at timestamptz not null default now(),

    confirmed_by text,

    confirmed_at timestamptz,

    created_at timestamptz not null default now(),

    constraint images_status_check
        check (status in ('pending', 'confirmed', 'rejected')),

    constraint image_source_check
        check (image_source in ('upload','url'))
);


/* Adding a notes table to allow comments to be made across
multiple fields e.g. notes on the abbreviation itself, the 
expansion(s), the images or the source*/
create table public.notes (

    id bigint generated always as identity primary key,

    entity_type text not null,

    entity_uuid uuid,

    entity_id bigint,

    source_name text,

    note text not null,

    created_at timestamptz not null default now(),

    created_by text,

    constraint entity_type_check
        check (entity_type in (
            'source',
            'abbreviation',
            'expansion',
            'image',
            'image_expansion'
        )),

    constraint entity_reference_check
        check (

            entity_type = 'source'
            and source_name is not null

            or

            entity_type != 'source'
            and (
                (entity_uuid is not null and entity_id is null)
                or
                (entity_uuid is null and entity_id is not null)
            )
        )
);