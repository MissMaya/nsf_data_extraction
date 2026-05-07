create table public.abbreviation_expansion (

    abbreviation_uuid uuid not null
        references public.abbreviations(uuid)
        on delete cascade,

    expansion_id bigint not null
        references public.expansions(id)
        on delete cascade,

    notes text,

    created_at timestamptz not null default now(),

    primary key (abbreviation_uuid, expansion_id)
);


create table public.image_expansion (

    image_id bigint not null
        references public.images(id)
        on delete cascade,

    expansion_id bigint not null
        references public.expansions(id)
        on delete cascade,

    status text not null default 'pending',

    notes text,

    created_at timestamptz not null default now(),

    primary key (image_id, expansion_id),

    constraint image_expansion_status_check
        check (status in ('pending', 'confirmed', 'rejected'))
);