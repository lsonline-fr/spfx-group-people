import PeopleCard from "../models/PeopleCard";
import { PersonaSize } from "office-ui-fabric-react/lib/Persona";

export interface IGroupPeopleProps {
  title: string;
  displayTitle: boolean;
  size: PersonaSize;
  users: Array<PeopleCard>;
  hide: boolean;
}
