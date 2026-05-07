create index abbreviations_abbreviation_idx
    on public.abbreviations (abbreviation);

create index abbreviations_source_idx
    on public.abbreviations (data_source);

create index expansions_text_idx
    on public.expansions (expansion_text);

create index expansions_type_idx
    on public.expansions (expansion_type);

create index images_abbreviation_uuid_idx
    on public.images (abbreviation_uuid);

create index images_status_idx
    on public.images (status);

create index image_expansion_status_idx
    on public.image_expansion (status);